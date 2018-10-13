package test;

import utils.JacobExcelTool;

public class Test {

	public static void main(String[] args) {
		JacobExcelTool tool = new JacobExcelTool();
		//打开
		//1、excel后缀为xlsm，为启用宏的工作簿
		//2、避免弹出警告框，操作步骤：文件-选项-信任中心-信任中心设置-隐私选项中把“保存时从文件属性中删除个人信息”前的勾号去掉
		tool.OpenExcel(System.getProperty("user.dir").replace("\\", "/") + "/src/test/JacobTest.xlsm",false,false);
		//调用Excel宏
		tool.callMacro("createPicture");
		tool.callMacro("createPicture2");
		//关闭并保存，释放对象
		tool.CloseExcel(true, true);
		System.out.println("调用成功！");
	}
}
