import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.http.HttpClient;
import java.net.http.HttpResponse;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class UploadFile{
    public void uploadFile(String url, String paramater, File file)
    {
        HttpClient httpclient = new DefaultHttpClient();
        httpclient.getParams().setParameter(CoreProtocolPNames.PROTOCOL_VERSION, HttpVersion.HTTP_1_1);

        HttpPost httppost = new HttpPost(url);
        MultipartEntity mpEntity = new MultipartEntity();

        ContentBody cbFile = new FileBody(file, "xlsx");
        mpEntity.addPart(parameter, cbFile);
        httppost.setEntity(mpEntity);
        HttpResponse response = httpclient.execute(httppost);
        HttpEntity resEntity = response.getEntity();

        String response1 = response.getStatusLine();
        httpclient.getConnectionManager().shutdown();

        System.out.println("Response: " + response1);
        //return response;
    }
}
