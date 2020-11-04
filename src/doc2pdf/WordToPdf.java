package doc2pdf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Scanner;
import com.aspose.words.Document;
import com.aspose.words.License;

public class WordToPdf {

	public static void main(String[] args) {
		System.err.println("路径格式：E:/XXX/XXX.docx");
		System.out.println("请输入文件或文件夹路径：");
		@SuppressWarnings("resource")
		Scanner scanner = new Scanner(System.in);
		String path = scanner.nextLine();
		if(StringUtils.isNullOrEmpty(path)){
			System.out.println("请输入文件路径！");
			return;
		}
		System.out.println(path);
		doc2pdf(path);
	}

	private static boolean getLicense() {
		boolean result = false;
		try {
			InputStream is = WordToPdf.class.getClassLoader().getResourceAsStream("license.xml"); // license.xml应放在..\WebRoot\WEB-INF\classes路径下
			License aposeLic = new License();
			aposeLic.setLicense(is);
			result = true;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}


	/**
	 * 同级目录下word文档转化为pdf
	 * @param file
	 *            文件名称（目录或文件）
	 *
	 */
	public static void doc2pdf(File file) {
		if (!getLicense()) { // 验证License 若不验证则转化出的pdf文档会有水印产生
			return;
		}
		if (!file.exists()) {
			return;
		}
		try {
			if (file.isDirectory()) {
				File[] files = file.listFiles();
				for (File fileObj : files) {
					doc2pdf(fileObj);
				}
			} else if(FileUtils.isWord(file.getName())){
				String pdfPath = file.getAbsolutePath();
				// 新建pdf文档名称
				String pdfName = file.getName().substring(0, file.getName().lastIndexOf(".")) + ".pdf";
				File pdfFile = new File(FileUtils.getPeerPath(pdfPath) + "\\" + pdfName); // 新建一个pdf文档
				FileOutputStream os = new FileOutputStream(pdfFile);
				Document doc = new Document(pdfPath); // Address是将要被转化的word文档
				doc.save(os, com.aspose.words.SaveFormat.PDF);// 全面支持DOC, DOCX,OOXML, RTF HTML,OpenDocument,PDF, EPUB, XPS, SWF 相互转换
				// 转换完成删除word
				new File(pdfPath).delete();
				os.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}


	public static void doc2pdf(String srcPath) {
		File srcFile = new File(srcPath);
		if (!srcFile.exists()){
			System.out.println("文件路径不存在！");
			return ;
		}
		// 同级目录
		String peerPath = FileUtils.getPeerPath(srcPath);
		//将源文件夹拷贝至同级目录下该文件夹
		String peerName = peerPath + "/" + srcFile.getName() + "_副本";
		File destFile = new File(peerName);
		// 若为目录
		if (srcFile.isDirectory()) {
			destFile.mkdir();
			try {
				// 拷贝目录
				FileUtils.copyDir(srcFile, destFile);
				File[] files = destFile.listFiles();
				for (File file : files) {
					doc2pdf(file);
				}

			} catch (IOException e) {
				e.printStackTrace();
			}
		} else {
			doc2pdf(srcFile);
		}
		System.out.println("转换完成！");
	}



}
