from dotenv import load_dotenv
import os

# 先加载环境变量
load_dotenv()

# 再读取
print(os.getenv("DASHSCOPE_API_KEY"))  # 现在应该输出你的密钥