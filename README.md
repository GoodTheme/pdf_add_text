# pdf_add_text
给pdf逐页加字。

在每页以(x, y)为文字框左上角座标添加文字，(0, 0)为页面左下角，(100, 100)为页面右上角；

使用/A替代逐页递增的自然数，在/A后紧跟着的<>中使用/B、/C和数字分别表示按书签递增、中文数字和起始数字（优先级高），例如“第/A</C10>页”表示从“第十页”开始；

可使用<br/>换行，<strike>删除线</strike>，<u>下划线</u><font color=red size=20>文字颜色、大小</font>等XML标记，少部分英文字体可用<b>加粗</b>，<i>斜体</i>。
