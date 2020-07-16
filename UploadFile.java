import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public String uploadFile(String url, String paramater, File file)
{
    HttpClient httpclient = new DefaultHttpClient();
    httpclient.getParams().setParameter(CoreProtocolPNames.PROTOCOL_VERSION, HttpVersion.HTTP_1_1);
    HttpPost httppost = new HttpPost(url);
    MultipartEntity mpEntity = new MultipartEntity();
    ContentBody cbFile = new FileBody(file, "image/jpeg");
    mpEntity.addPart(parameter, cbFile);
    httppost.setEntity(mpEntity);
    HttpResponse response = httpclient.execute(httppost);
    HttpEntity resEntity = response.getEntity();
    String response = response.getStatusLine();
    httpclient.getConnectionManager().shutdown();
    return response;
}
