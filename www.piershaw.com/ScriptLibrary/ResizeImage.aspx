<%@ Page Language="C#"%>
<%@ Import Namespace="System.Drawing" %>
<%@ Import Namespace="System.Drawing.Imaging" %>
<script runat="server">
// Smart Image Processor 1.1
// Version: 1.1.2

	void Page_Load(Object s, EventArgs e) {
		
		int intNewWidth,intNewHeight, maxWidth = 10000, maxHeight = 10000, qQuality = 100;
    if ( Request["w"] != null) maxWidth = int.Parse(Request["w"]);
    if ( Request["h"] != null) maxHeight = int.Parse(Request["h"]);
    if ( Request["q"] != null) qQuality = int.Parse(Request["q"]);
		
		//get image from parameter
		string pictureFileName = Request["f"];
    string newFileName = Request["nf"];
    if (pictureFileName == null || pictureFileName == "" || !System.IO.File.Exists(pictureFileName)) {
      Response.Write("Error: File (" + pictureFileName + ") not found or empty");  
      return;
    }
		System.Drawing.Image inputImage = System.Drawing.Image.FromFile(pictureFileName);
        
    //define size for new image
		string aspect = Request["a"];
		if (aspect == "true") {
			if (maxWidth < inputImage.Width || maxHeight < inputImage.Height) {
				if (maxWidth >= maxHeight) {
					intNewWidth = (int)((double)maxHeight*((double)inputImage.Width/(double)inputImage.Height));
					intNewHeight = maxHeight;
				} else {
					intNewWidth = maxWidth;
					intNewHeight = (int)((double)maxWidth*((double)inputImage.Height/(double)inputImage.Width));
				}
				if (intNewWidth > maxWidth) {
					intNewWidth = maxWidth;
					intNewHeight = (int)((double)maxWidth*((double)inputImage.Height/(double)inputImage.Width));
				}
				if (intNewHeight > maxHeight) {
					intNewWidth = (int)((double)maxHeight*((double)inputImage.Width/(double)inputImage.Height));
					intNewHeight = maxHeight;
				}
			} else {
				intNewWidth = inputImage.Width;
				intNewHeight = inputImage.Height;
			}
		} else {
				intNewWidth = maxWidth;
				intNewHeight = maxHeight;
		}

    try {        
      //output new image with different size
  		Bitmap outputBitMap = new Bitmap(inputImage,intNewWidth,intNewHeight);
      inputImage.Dispose();
     	EncoderParameters eps = new System.Drawing.Imaging.EncoderParameters(1);
     	eps.Param[0] = new System.Drawing.Imaging.EncoderParameter( System.Drawing.Imaging.Encoder.Quality, qQuality );
     	ImageCodecInfo ici = GetEncoderInfo("image/jpeg");
      if (pictureFileName.ToLower() == newFileName.ToLower())
        System.IO.File.Delete(pictureFileName);
     	outputBitMap.Save( newFileName, ici, eps );
      outputBitMap.Dispose();      
    }		
    catch (Exception ex) {
    	Response.Write("Error: " + ex);
			return;
    }  
    
    Response.Write(intNewWidth + ";" + intNewHeight + ";" + "DONE");
  }
    
  private static ImageCodecInfo GetEncoderInfo(String mimeType) {
    int j;
    ImageCodecInfo[] encoders;
    encoders = ImageCodecInfo.GetImageEncoders();
    for(j = 0; j < encoders.Length; ++j) {
      if(encoders[j].MimeType == mimeType)
        return encoders[j];
    }
    return null;
  }
    
</script>
