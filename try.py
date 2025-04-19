# 导入必要库
import pandas as pd  # 核心数据分析库
import os  # 操作系统接口
import re  # 正则表达式处理
import random  # 随机数生成
from openpyxl import load_workbook  # Excel文件读写支持[1](@ref)
from openai import OpenAI  # OpenAI API客户端

class ExcelAIProcessor:
    def __init__(self):
        """初始化处理器核心组件"""
        self.df = None  # 存储Excel数据的DataFrame
        self.client = OpenAI(  # 配置安全API客户端
            api_key=os.getenv("DASHSCOPE_API_KEY"),  # 从环境变量获取API密钥
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"  # 阿里云兼容端点
        )
        self.safe_globals = {  # 创建沙箱执行环境
            'pd': pd,  # 仅允许访问pandas
            'df': None,  # 数据容器占位符
            'random': random,  # 受控随机数模块
            '__builtins__': {  # 限制内置函数白名单
                'str': str, 'int': int, 'float': float, 'bool': bool,
                'list': list, 'dict': dict, 'tuple': tuple,
                'len': len, 'range': range  # 基础数据操作函数
            }
        }

    def read_excel(self, file_path):
        """读取并预处理Excel文件"""
        try:
            self.df = pd.read_excel(file_path)  # 使用pandas读取Excel
            # 列名标准化处理（替换特殊字符为下划线）
            original_columns = self.df.columns.tolist()
            self.df.columns = self.df.columns.str.replace(r'[^\w]', '_', regex=True)
            # 打印列名映射关系
            print("列名标准化映射：")
            for o, n in zip(original_columns, self.df.columns):
                print(f"原始: {o} → 新: {n}")
            # 输出数据概况
            print(f"\n成功读取Excel文件，共{len(self.df)}行{len(self.df.columns)}列")
            print("前3行数据样例：")
            print(self.df.head(3))  # 展示数据前3行
            return True
        except Exception as e:
            print(f"读取文件失败: {str(e)}")  # 异常处理
            return False

    def generate_pandas_code(self, instruction):
        """通过AI生成安全的Pandas代码"""
        # 构造包含数据样例的prompt
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
            # 调用大模型API生成代码
            completion = self.client.chat.completions.create(
                model="deepseek-r1",  # 指定模型版本
                messages=[{"role": "user", "content": prompt}],
                temperature=0.2  # 低随机性保证代码稳定性
            )
            content = completion.choices[0].message.content
            # 使用正则提取代码块
            code_match = re.search(r'<code>(.*?)</code>', content, re.DOTALL)
            return self._clean_code(code_match.group(1).strip()) if code_match else None
        except Exception as e:
            print(f"API请求失败: {str(e)}")
            return None

    def _clean_code(self, code):
        """代码安全清洗"""
        # 移除注释和空行
        code = ' '.join([line.split('#')[0].strip() for line in code.split('\n') 
                        if line.strip() and not line.startswith(('import', 'from'))])
        
        # 检测危险模式（如系统调用）
        unsafe_patterns = r'(numpy|__import__|os\.|sys\.|open\(|eval\()'
        if re.search(unsafe_patterns, code, re.IGNORECASE):
            print("检测到危险操作或禁用库")
            return None
            
        # 验证是否包含df操作
        if not re.search(r'\bdf\b', code):
            print("代码未包含df操作")
            return None
            
        return code

    def safe_execute(self, code):
        """在沙箱环境中执行生成的代码"""
        if not code:
            return False
            
        try:
            local_vars = {'df': self.df.copy(deep=True)}  # 创建数据副本防止污染
            # 编译执行代码
            exec(compile(code, '<string>', 'exec'), self.safe_globals, local_vars)
            
            # 验证执行结果
            new_df = local_vars.get('df')
            if not isinstance(new_df, pd.DataFrame) or new_df.empty:
                print("无效的DataFrame结果")
                return False
                
            # 保持原始数据类型
            common_cols = new_df.columns.intersection(self.df.columns)
            try:
                new_df[common_cols] = new_df[common_cols].astype(self.df.dtypes[common_cols].to_dict())
            except Exception as e:
                print(f"类型转换警告: {str(e)}")
                
            self.df = new_df  # 更新主数据集
            print(f"执行成功，更新后数据：\n{self.df.head(3)}")
            return True
        except Exception as e:
            print(f"执行失败: {str(e)}")
            return False

    def save_excel(self, output_path):
        """保存处理结果到Excel"""
        try:
            self.df.to_excel(output_path, index=False, engine='openpyxl')  # 使用openpyxl引擎
            print(f"文件已保存到 {output_path}")
            return True
        except Exception as e:
            print(f"保存失败: {str(e)}")
            return False

if __name__ == "__main__":
    # 创建处理器实例
    processor = ExcelAIProcessor()
    
    # 获取输入文件路径
    input_file = input("请输入Excel文件路径: ").strip()
    if not processor.read_excel(input_file):
        exit()

    # 交互式指令处理循环
    while True:
        instruction = input("\n操作指令（输入'save'保存）: ").strip()
        if instruction.lower() == 'save':
            break
            
        if code := processor.generate_pandas_code(instruction):
            print(f"生成代码: {code}")
            processor.safe_execute(code) or print("失败建议：请使用标准化后的列名")
        else:
            print("代码生成失败")

    # 保存处理结果
    output_file = input("保存路径（直接回车覆盖原文件）: ").strip() or input_file
    processor.save_excel(output_file)