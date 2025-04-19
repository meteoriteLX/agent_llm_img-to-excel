import pandas as pd
import os
import re
from openai import OpenAI

class ExcelAIProcessor:
    def __init__(self):
        self.df = None
        self.client = OpenAI(
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )
        self.safe_globals = {'pd': pd, 'df': None}

    def read_excel(self, file_path):
        """读取Excel文件到DataFrame"""
        try:
            self.df = pd.read_excel(file_path)
            print(f"成功读取Excel文件，共{len(self.df)}行{len(self.df.columns)}列")
            return True
        except Exception as e:
            print(f"读取文件失败: {str(e)}")
            return False

    def generate_pandas_code(self, instruction):
        """使用DeepSeek API生成Pandas操作代码"""
        # 构建包含数据结构的提示词
        prompt = f"""当前DataFrame结构：
{self.df.head(3).to_markdown(index=False)}

请将以下自然语言指令转换为安全的Pandas代码：
指令：{instruction}

要求：
1. 只生成一行有效的Python代码
2. 使用df变量名操作DataFrame
3. 禁止文件操作和其他危险操作
4. 返回格式：<code>你的代码</code>"""

        try:
            completion = self.client.chat.completions.create(
                model="deepseek-r1",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2
            )
            
            # 提取生成的代码
            content = completion.choices[0].message.content
            code_match = re.search(r'<code>(.*?)</code>', content, re.DOTALL)
            return code_match.group(1).strip() if code_match else None
        except Exception as e:
            print(f"API请求失败: {str(e)}")
            return None

    def safe_execute(self, code):
        """安全执行生成的代码"""
        if not code:
            return False
            
        try:
            # 限制执行环境并创建数据副本
            local_vars = {'df': self.df.copy()}
            byte_code = compile(code, '<string>', 'exec')
            exec(byte_code, {'pd': pd, '__builtins__': {}}, local_vars)
            
            # 验证结果有效性
            new_df = local_vars.get('df', self.df)
            if isinstance(new_df, pd.DataFrame) and not new_df.empty:
                self.df = new_df
                print("执行成功，DataFrame已更新")
                return True
            return False
        except Exception as e:
            print(f"代码执行失败: {str(e)}")
            return False

    def save_excel(self, output_path):
        """保存修改后的DataFrame到Excel"""
        try:
            self.df.to_excel(output_path, index=False)
            print(f"文件已保存到 {output_path}")
            return True
        except Exception as e:
            print(f"保存失败: {str(e)}")
            return False

if __name__ == "__main__":
    processor = ExcelAIProcessor()
    
    # 输入文件路径
    input_file = input("请输入要处理的Excel文件路径: ")
    if not processor.read_excel(input_file):
        exit()

    # 处理指令循环
    while True:
        instruction = input("\n请输入操作指令（输入'save'保存并退出）: ")
        if instruction.lower() == 'save':
            break
        
        code = processor.generate_pandas_code(instruction)
        if code:
            print(f"生成的代码: {code}")
            if processor.safe_execute(code):
                print(processor.df.head(3))  # 显示修改后的前3行
            else:
                print("执行失败，请尝试重新描述需求")
        else:
            print("代码生成失败")

    # 保存文件
    output_file = input("请输入保存路径（默认按原文件保存）: ") or input_file
    processor.save_excel(output_file)