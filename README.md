# 微信聊天记录分析
## 灵感来源于<a href="https://github.com/godweiyang/wechat-explore">Wechat Explore</a>
## 功能
1. 导出微信数据库csv文件，可使用<a href="https://github.com/godweiyang/wechat-explore">Wechat Explore</a>和<a href="https://github.com/godweiyang/wechat-analysis">Wechat Analysis</a>进一步处理
2. 自动提取信息并导出为xlsx
3. 将聊天记录按人或群聊导出为单独文件
## 使用方法（仅支持Windows系统）
1. 在Root过的手机或模拟器上导入要分析的聊天记录
2. 提取出/data/data/com.tencent.mm/shared_prefs/auth_info_key_prefs.xml和/data/data/com.tencent.mm/MicroMsg/长串字母文件夹/EnMicroMsg.db
3. 安装Python
4. 安装库pandas,openpyxl
5. 运行main.py并按提示填入两个文件的路径
6. 导出的message.csv可进一步处理，message.xlsx为处理过的聊天记录，chats文件夹下为按人或群单独的聊天记录，带chatroom的为群聊聊天记录。
## 已知问题
1. 人或群名只支持微信号，且微信号为初始微信号，可能要根据聊天记录判断对话人，目前无解，了解解决方案的欢迎提出来