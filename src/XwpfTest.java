import com.sun.prism.shader.Solid_TextureYV12_AlphaTest_Loader;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.commons.io.FileUtils;
import org.junit.Test;
import java.io.*;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by zhanghao on 16/1/8.
 */

public class XwpfTest {
    /**
     * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
     * @throws Exception
     */
    @Test
    public void testTemplateWrite() throws Exception
    {
        Map<String,String> params = new HashMap<>();
        params.put("time","2016年01月12日");
        params.put("money"," 11.2 ");
        params.put("zheng","壹拾壹万贰仟圆整");
        params.put("company", " 你好吗公司 ");
        params.put("xiaoshi","2h");
        params.put("daxie", " 两小时 ");
        String inputfilePath = "/Users/Sunbelife/Desktop/自动化合同/hetong_v2.docx";
        System.out.println("来源路径："+ inputfilePath);
        System.out.println("匹配中");
        InputStream is = new FileInputStream(inputfilePath);//打开输入流
        XWPFDocument doc = new XWPFDocument(is);
        //替换段落里面的变量
//      this.replacedoc(doc,params);
        this.replaceInPara(doc, params);
        this.replaceInTable(doc, params);
        //输出至 txt
        this.exportxt(doc,params);
        //替换表格里面的变量
        String outputfilepath = "/Users/Sunbelife/Desktop/自动化合同/hetong_v3.docx";
        FileUtils.deleteQuietly(new File(outputfilepath));//更新
        OutputStream os = new FileOutputStream(outputfilepath);
        System.out.println("Docx 格式版本已输出至："+ outputfilepath );
        doc.write(os);
        this.close(os);
        this.close(is);
    }

    /**
     * 替换段落里面的变量
     * @param doc 要替换的文档
     * @param params 参数
     */

    private void replaceInPara(XWPFDocument doc, Map<String, String> params)
    {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext())
        {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    /**
     * 替换段落里面的变量
     * @param para 要替换的段落
     * @param params 参数
     */

    private void replaceInPara(XWPFParagraph para, Map<String,String> params)
    {
        List<XWPFRun> runs;
        Matcher matcher;
        if (this.matcher(para.getParagraphText()).find())
        {
            runs = para.getRuns();
            if (this.matcher(para.getParagraphText()).find())
            System.out.print("匹配到含有需要替换的字符段落：" + para.getText() + "替换为");
            for (int i=0; i<runs.size(); i++)
            {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                matcher = this.matcher(runText);
                if (matcher.find())
                {
                    while ((matcher = this.matcher(runText)).find())
                    {
                        runText = matcher.replaceFirst(String.valueOf(params.get(matcher.group(1))));
                    }
                    //找到之后直接覆盖值
                    runs.get(i).setText(runText, 0);
                    System.out.println(runText);
                }
            }
        }
    }

    /**
     * 替换表格
     *
     */

    private void replaceInTable(XWPFDocument doc, Map<String,String> params) {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        this.replaceInPara(para, params);
                    }
                }
            }
        }
    }

    private void exportxt(XWPFDocument doc,Map<String, String> params)
    {
        XWPFWordExtractor extractor = new XWPFWordExtractor(doc);
        String txtpath = "/Users/Sunbelife/Desktop/自动化合同/输出的 doc.txt";
        FileUtils.deleteQuietly(new File(txtpath)); //先删除
        String Text = extractor.getText();
        //替换文本
//        for (String key : params.keySet())
//        {
//            System.out.println("将"+ key + "替换为" + params.get(key));
//            Text = Text.replace(key, params.get(key)); 替换文本
//        }
        try
        {
            FileUtils.write(new File(txtpath),Text,"UTF-8",true); //再写入
        } catch (IOException e)
        {
            e.printStackTrace();
        }
        System.out.println("TXT 格式版本已输出至" + txtpath);
    }

    /**
     * 关闭输入流
     * @param is
     */

    private void close(InputStream is)
    {
        if (is != null)
        {
            try
            {
                is.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }


    private Matcher matcher(String str)
    {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 关闭输出流
     * @param os
     */

    private void close(OutputStream os)
    {
        if (os != null)
        {
            try
            {
                os.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }



}
