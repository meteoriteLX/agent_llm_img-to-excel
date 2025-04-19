import pandas as pd
import os
import re
import random
from openai import OpenAI

class ExcelAIProcessor:
    def __init__(self):
        self.df = None
        self.client = OpenAI(
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        self.safe_globals = {
            'pd': pd,
            'df': None,
            'random': random,
            '__builtins__': {
                'str': str, 'int': int, 'float': float, 'bool': bool,
                'list': list, 'dict': dict, 'tuple': tuple,
                'len': len, 'range': range
            }
        }

    def read_excel(self, file_path):
        try:
            self.df = pd.read_excel(file_path)
            # 标准化列名并打印
            original_columns = self.df.columns.tolist()
            self.df.columns = self.df.columns.str.replace(r'[^\w]', '_', regex=True)
            print("列名标准化映射：")
            for o, n in zip(original_columns, self.df.columns):
                print(f"原始: {o} → 新: {n}")
            print(f"\n成功读取Excel文件，共{len(self.df)}行{len(self.df.columns)}列")
            print("前3行数据样例：")
            print(self.df.head(3))
            return True
        except Exception as e:
            print(f"读取文件失败: {str(e)}")
            return False

    def generate_pandas_code(self, instruction):
        prompt = f"""当前DataFrame结构（显示前3行）：
{self.df.head(3).to_markdown(index=False)}

当前标准列名列表：{', '.join(self.df.columns.tolist())}

请将以下自然语言指令转换为安全的Pandas代码：
指令：{instruction}

要求：
1. 必须使用上述标准列名列表中的准确列名
2. 生成单行Python代码，使用df变量操作DataFrame
3. 删除列操作必须明确指定列名，格式：df.drop(columns=[列名1, 列名2])
4. 兼容Pandas >=1.0版本语法
5. 如果要生成随机数，则必须使用random模块生成随机值（如random.randint）
6. 禁止直接import random模块（已预注入）
7. 返回格式：<code>你的代码</code>"""

        try:
            completion = self.client.chat.completions.create(
                model="deepseek-r1",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2
            )
            content = completion.choices[0].message.content
            code_match = re.search(r'<code>(.*?)</code>', content, re.DOTALL)
            return self._clean_code(code_match.group(1).strip()) if code_match else None
        except Exception as e:
            print(f"API请求失败: {str(e)}")
            return None

    def _clean_code(self, code):
        code = ' '.join([line.split('#')[0].strip() for line in code.split('\n') 
                        if line.strip() and not line.startswith(('import', 'from'))])
        
        unsafe_patterns = r'(numpy|__import__|os\.|sys\.|open\(|eval\()'
        if re.search(unsafe_patterns, code, re.IGNORECASE):
            print("检测到危险操作或禁用库")
            return None
            
        if not re.search(r'\bdf\b', code):
            print("代码未包含df操作")
            return None
            
        return code

    def safe_execute(self, code):
        if not code:
            return False
            
        try:
            local_vars = {'df': self.df.copy(deep=True)}
            exec(compile(code, '<string>', 'exec'), self.safe_globals, local_vars)
            
            new_df = local_vars.get('df')
            if not isinstance(new_df, pd.DataFrame) or new_df.empty:
                print("无效的DataFrame结果")
                return False
                
            # 修正类型检查逻辑
            common_cols = new_df.columns.intersection(self.df.columns)
            try:
                new_df[common_cols] = new_df[common_cols].astype(self.df.dtypes[common_cols].to_dict())
            except Exception as e:
                print(f"类型转换警告: {str(e)}")
                
            self.df = new_df
            print(f"执行成功，更新后数据：\n{self.df.head(3)}")
            return True
        except Exception as e:
            print(f"执行失败: {str(e)}")
            return False

    def save_excel(self, output_path):
        try:
            self.df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"文件已保存到 {output_path}")
            return True
        except Exception as e:
            print(f"保存失败: {str(e)}")
            return False

if __name__ == "__main__":
    processor = ExcelAIProcessor()
    
    input_file = input("请输入Excel文件路径: ").strip()
    if not processor.read_excel(input_file):
        exit()

    while True:
        instruction = input("\n操作指令（输入'save'保存）: ").strip()
        if instruction.lower() == 'save':
            break
            
        if code := processor.generate_pandas_code(instruction):
            print(f"生成代码: {code}")
            processor.safe_execute(code) or print("失败建议：请使用标准化后的列名")
        else:
            print("代码生成失败")

    output_file = input("保存路径（直接回车覆盖原文件）: ").strip() or input_file
    processor.save_excel(output_file)