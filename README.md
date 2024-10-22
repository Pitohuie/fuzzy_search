

### 使用说明：模糊搜索姓名的小程序

#### 1. 下载并安装

1. **获取 EXE 文件**: 将提供dist目录下的[SES.zip](https://github.com/Pitohuie/fuzzy_search/releases/download/v0.1.0/SES.zip)


下载到你的电脑。你可以将它放置在一个易于访问的文件夹中，如桌面或指定的项目目录中。

2. **双击运行**: 双击 EXE 文件以启动应用程序。你会看到一个简洁的用户界面，提供文件选择、列名输入、关键词输入和结果显示区域。

#### 2. 用户界面说明

- **选择Excel文件**: 点击此按钮以选择包含你要搜索的姓名的 Excel 文件。
  
- **请输入第一个列名**: 在这个输入框中输入 Excel 表格中要搜索的第一个列名（例如 "Name" 或 "姓名"）。

- **请输入第二个列名**: 你可以在这里输入第二个列名，如果需要在另一个列中进行模糊搜索。

- **请输入第三个列名**: 你可以在这里输入第三个列名，增加搜索的灵活性。

- **请输入要搜索的关键词**: 在这里输入你想要模糊匹配的关键词（例如一个姓名的部分拼写或全拼）。

- **搜索按钮**: 点击此按钮开始搜索。程序将自动在你指定的列中进行模糊匹配，并显示最相似的结果。

- **结果显示区域**: 搜索结果将在这里显示，包括匹配的内容及其相似度。

#### 3. 详细步骤

1. **选择文件**: 启动应用程序后，首先点击“选择Excel文件”按钮，找到并打开你要搜索的 Excel 文件。

2. **输入列名**: 在“请输入列名”输入框中输入要搜索的列名。如果你的 Excel 表格中姓名所在的列名为 "Name"，请在输入框中输入 "Name"。

3. **输入关键词**: 在“请输入要搜索的关键词”输入框中，输入你想搜索的姓名或关键词。

4. **点击搜索**: 输入完毕后，点击“搜索”按钮。程序会根据你的输入，在指定的列中寻找最相似的匹配，并将结果显示在下方的区域。

5. **查看结果**: 在结果显示区域，你会看到程序找到的最相似的结果列表，以及每个结果的相似度百分比。

#### 4. 常见问题解答

1. **找不到列**:
   - 确保你输入的列名与 Excel 表格中的列名完全一致，包括大小写。

2. **没有匹配结果**:
   - 如果没有找到匹配的结果，可能是因为关键词与列中内容的相似度过低。尝试输入更准确或更多的关键词。

3. **程序崩溃或无响应**:
   - 如果程序意外关闭或无响应，请确保你的 Excel 文件格式正确（例如，.xlsx 格式）并且没有被其他程序占用。

#### 5. 退出程序

- 当你完成搜索后，可以直接关闭窗口来退出程序。

### 系统要求

- **操作系统**: Windows 10 或更高版本。
- **Python 环境**: 不需要。EXE 文件已经打包所有必要的依赖项，可以独立运行。

### 注意事项

- **文件安全性**: 确保你从信任的来源获取 EXE 文件，以避免可能的安全风险。
- **备份数据**: 在操作重要数据时，建议备份 Excel 文件，以防意外修改或损坏。

### 联系支持

- 如果你遇到问题或有功能建议，请联系程序开发者提供的支持渠道。
- QQ：1748492875
