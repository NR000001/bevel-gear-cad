# bevel-gear-cad

## SolidWorks VBA 宏说明

已新增宏：`macros/ExportDrawing.bas`。

### 宏功能
1. 打开 `source/从动锥齿轮.SLDPRT`（只读，不修改原始零件）。
2. 新建 A3 横向工程图。
3. 插入主视图、俯视图、右视图、等轴测视图。
4. 添加参数说明文字（可在代码中修改）。
5. 导出以下文件到模型所在目录：
   - `从动锥齿轮工程图.SLDDRW`
   - `从动锥齿轮工程图.pdf`
   - `从动锥齿轮工程图.dxf`
   - `从动锥齿轮.step`

### 使用方法
1. 用 SolidWorks 打开宏：`macros/ExportDrawing.bas`。
2. 运行 `main` 过程。
3. 确保 SolidWorks 已配置默认工程图模板（Drawing Template），否则无法自动创建工程图。

### 目录说明
- `source/从动锥齿轮.SLDPRT`：指向根目录零件文件的链接，用于满足固定路径调用。
- `从动锥齿轮.SLDPRT`：原始零件文件（未被宏修改）。
