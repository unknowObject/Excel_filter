# Excel_filter
A filter tool which helps analyze .xls file according to the key words typed in the config files.

使用说明：

  工具目前支持三种筛选模式，分别为：包含，包含（与），包含（或）。每种筛选模式对应一个配置文件，包含对应“config_contain.txt”，包含（与）对应“config_and.txt”，包含（或）对应“config_or.txt”。使用前需检查对应配置文件是否有误。

  使用时，先选择要分析的文件（.xls格式）或直接输入路径，再选择输出文件目标路径。输出文件名为"result.xls"。再在界面左下角输入需要分析的内容所在表格中的列数以及在下拉菜单中选择筛选模式，表格名需为"Sheet1"。
  
  设置完成后点击开始分类即可针对输入表格应用所选筛选模式进行分类并将分类结果输出至目标路径下的"result.xls"文件中，多次分类后工具会将原有"result.xls"文件覆盖。"result.xls"文件第一行为占位行替换一般表格文件中表头的位置。分类结果将作为一列表格输出在文件中。


配置文件格式需求：

  "config_contain.txt"中，每行由关键词和关键词所属类别构成，前者为关键词，后者为类别，中间以tab隔开，输入类别完成后立刻使用回车换行。（注：从tab开始直至回车换行为止所有内容都会被识别为类别，关键词同理，从每行开始至tab前所有内容都会被识别为关键词）

  最后一行内容输入完成后也要再用回车进行一次换行，如：“
  
关键词x  类别y

”

  "config_and.txt"以及"config_or.txt"中，与"config_contain.txt"类似，前两个词为关键词，以tab隔开，后一个词为所属类别，同样以tab隔开，输入完成后立刻使用回车换行。最后一行输入后同样需要用一次回车换行。
