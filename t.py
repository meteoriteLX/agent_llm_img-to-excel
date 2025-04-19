import requests

# DeepSeek 官方配置
headers = {"Authorization": "Bearer sk-c030f7b47ca045a2ae44978b09092298"}
payload = {
    "model": "deepseek-r1-distill-qwen-1.5b",  # 正确模型名称
    "messages": [{"role": "user", "content": "test"}]
}
response = requests.post(
    "https://dashscope.aliyuncs.com/compatible-mode/v1",  # DeepSeek官方终端节点
    json=payload,
    headers=headers
)
print(response.status_code, response.text)