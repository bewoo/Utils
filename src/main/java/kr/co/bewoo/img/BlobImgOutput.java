package kr.co.bewoo.img;

import java.io.IOException;
import java.io.OutputStream;

import javax.servlet.http.HttpServletResponse;

public class BlobImgOutput {
    public static void blobImgOutput(HttpServletResponse response) throws IOException{
        byte[] image = new byte[]{1,2,3,4,5};   //dummy data
        response.setContentType("image/gif");
        response.setContentLength(image.length);
        OutputStream os = response.getOutputStream();
        os.write(image);
        os.flush();
    }
}
