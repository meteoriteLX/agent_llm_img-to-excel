import pandas as pd
import os
import re
import random  # 添加内置随机模块
from openai import OpenAI

class ExcelAIProcessor:
    def __init__(self):
        self.df = None
        self.client = OpenAI(
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        # 增强安全白名单
        self.safe_globals = {
            'pd': pd,
            'df': None,
            'random': random,  # 注入随机模块
            '__builtins__': {
                'str': str,
                'int': int,
                'float': float,
                'bool': bool,
                'list': list,
                'dict': dict,
                'tuple': tuple,
                'len': len,
                'range': range
            }
        }

    def read_excel(self, file_path):
        """读取Excel文件到DataFrame并标准化列名"""
        try:
            self.df = pd.read_excel(file_path)
            # 标准化列名：去除空格和特殊字符
            self.df.columns = self.df.columns.str.replace(r'[^\w]', '_', regex=True)
            print(f"成功读取Excel文件，共{len(self.df)}行{len(self.df.columns)}列")
            print("前3行数据样例：")
            print(self.df.head(3))
            return True
        except Exception as e:
            print(f"读取文件失败: {str(e)}")
            return False

    def generate_pandas_code(self, instruction):
        """使用DeepSeek API生成Pandas操作代码"""
        prompt = f"""当前DataFrame结构（显示前3行）：
{self.df.head(3).to_markdown(index=False)}

列名列表：{', '.join(self.df.columns.tolist())}

请将以下自然语言指令转换为安全的Pandas代码：
指令：{instruction}

要求：
1. 生成单行Python代码，使用df变量操作DataFrame
2. 如果要生成随机数，则必须使用random模块生成随机值（如random.randint）
3. 禁止直接import random模块（已预注入）
4. 兼容Pandas >=1.0版本语法
5. 禁止文件操作、网络请求和模块导入
6. 优先使用pandas原生方法
7. 返回格式：<code>你的代码</code>
8. 必须明确指定要删除的列名，禁止使用随机选择逻辑"""
        

        try:
            completion = self.client.chat.completions.create(
                model="deepseek-r1",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2
            )
            
            content = completion.choices[0].message.content
            code_match = re.search(r'<code>(.*?)</code>', content, re.DOTALL)
            if not code_match:
                return None
                
            code = code_match.group(1).strip()
            return self._clean_code(code)
        except Exception as e:
            print(f"API请求失败: {str(e)}")
            return None

    def _clean_code(self, code):
        """增强代码清洗逻辑"""
        # 移除注释和多余空行
        code_lines = []
        for line in code.split('\n'):
            line = line.split('#')[0].strip()
            if line and not line.startswith(('import', 'from')):
                code_lines.append(line)
        code = ' '.join(code_lines)
        
        # 增强安全检测
        unsafe_patterns = r'(?:os\.|sys\.|open\(|eval\(|__import__)'
        if re.search(unsafe_patterns, code):
            print("检测到危险操作，拒绝执行")
            return None
        
        if not re.search(r'\bdf\b', code):
            print("代码未包含df操作，拒绝执行")
            return None
            
        return code

    def safe_execute(self, code):
        """安全执行生成的代码"""
        if not code:
            return False
            
        try:
            local_vars = {'df': self.df.copy(deep=True)}
            byte_code = compile(code, '<string>', 'exec')
            exec(byte_code, self.safe_globals, local_vars)
            
            new_df = local_vars.get('df', None)
            # 增强结果验证
            if not isinstance(new_df, pd.DataFrame):
                print("执行结果不是有效的DataFrame")
                return False
                
            if len(new_df) == 0 and len(new_df.columns) == 0:
                print("执行后DataFrame完全为空，已拒绝修改")
                return False
                
            try:
                new_df.astype(self.df.dtypes.to_dict())
            except Exception as e:
                print(f"类型转换错误: {str(e)}")
                return False
                
            self.df = new_df
            print("执行成功，更新后的数据结构：")
            print(f"行数：{len(self.df)}，列数：{len(self.df.columns)}")
            print("前3行数据：")
            print(self.df.head(3))
            return True
        except SyntaxError as e:
            print(f"语法错误: {str(e)}")
            return False
        except Exception as e:  # 移除非法的PandasError捕获
            print(f"执行异常: {str(e)}")
            return False

    def save_excel(self, output_path):
        """保存修改后的DataFrame到Excel"""
        try:
            self.df.to_excel(
                output_path,
                index=False,
                engine='openpyxl'
            )
            print(f"文件已成功保存到 {output_path}")
            return True
        except PermissionError:
            print("文件保存失败：请关闭文件后重试")
            return False
        except Exception as e:
            print(f"保存失败: {str(e)}")
            return False

if __name__ == "__main__":
    processor = ExcelAIProcessor()
    
    input_file = input("请输入要处理的Excel文件路径: ").strip()
    if not processor.read_excel(input_file):
        exit()

    while True:
        instruction = input("\n请输入操作指令（输入'save'保存并退出）: ").strip()
        if instruction.lower() == 'save':
            break
        
        code = processor.generate_pandas_code(instruction)
        if code:
            print(f"生成的代码: {code}")
            if not processor.safe_execute(code):
                print("执行失败，请尝试以下解决方案：")
                print("1. 检查列名是否匹配当前数据")
                print("2. 使用更明确的指令（如指定随机范围）")
                print("3. 分步描述复杂操作")
        else:
            print("代码生成失败，请尝试更明确的指令")

    output_file = input(f"请输入保存路径（直接回车覆盖原文件）: ").strip() or input_file  # 修改默认路径为原文件
    processor.save_excel(output_file)