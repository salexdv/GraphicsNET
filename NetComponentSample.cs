using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Windows;
using System.ComponentModel;
using System.IO;
using System.Reflection;


namespace GraphicsNET
{
    public class ImageConverter : System.Windows.Forms.AxHost
    {
        public ImageConverter()
            : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public object ImageToIpicture(System.Drawing.Image image)
        {
            return ImageConverter.GetIPictureDispFromPicture(image);            
        }        
    }

    [ComVisible(true), Guid("8A603F7F-1DCB-4f2f-B929-27B9AD6E9BFA"), ProgId("AddIn.GraphicsNET")]
	public class NetComponentSampleCS : AddInLib.IInitDone, AddInLib.ILanguageExtender 
	{
        const string c_AddinName = "GraphicsNET";

        #region APImetafile
        public const uint CF_METAFILEPICT = 3;
        public const uint CF_ENHMETAFILE = 14;

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool CloseClipboard();

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetClipboardData(uint format);

        [DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern bool IsClipboardFormatAvailable(uint format);
        #endregion

        #region "IInitDone implementation"

        public NetComponentSampleCS()
        {                        
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void New()
        {            
        }

        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {            
            Data1C.Object1C = pConnection;
            Marshal.GetIUnknownForObject(Data1C.Object1C);            
        }

        public void Done()
        {
            Marshal.Release(Marshal.GetIDispatchForObject(Data1C.Object1C));
            Marshal.ReleaseComObject(Data1C.Object1C);
            Data1C.Object1C = null;
            �����������.Dispose();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void GetInfo(ref object[] pInfo)
        {            
            ((System.Array)pInfo).SetValue("2000", 0);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        #endregion
					  
		#region "����������"		
        public Image �����������;
        public object App1C = null;
        public object Pic1C = null;
        public object �������������� = null;
        private MemoryStream msBinary = null;        
        string ������������;
		#endregion
															   
		public void RegisterExtensionAs(ref string bstrExtensionName)
		{            
			bstrExtensionName = c_AddinName;            
		}
		
		#region "��������"
		enum Props
		{   //�������� �������������� ������� ����� ������� ����������
			ImageSize = 0,  
            Width = 1, 
            Height = 2,
            BinaryData = 3,
            Image = 4,
            IPicture = 5,
            LastProp = 6
		}
         
		public void GetNProps(ref int plProps)
		{	//����� 1� �������� ���������� ��������� �� �� �������
			plProps = (int)Props.LastProp;
		}
        
		public void FindProp(string bstrPropName, ref int plPropNum)
		{	//����� 1� ���� �������� ������������� �������� �� ��� ���������� �����
			switch(bstrPropName)
			{
				case "ImageSize": 
				case "�����������������":
					plPropNum = (int)Props.ImageSize;
					break;
                case "Height":
				case "������":
					plPropNum = (int)Props.Height;
					break;
                case "Width":
                case "������":
                    plPropNum = (int)Props.Width;
                    break;
                case "BinaryData":
                case "��������������":
                    plPropNum = (int)Props.BinaryData;
                    break;
                case "Image":
                case "�����������":
                    plPropNum = (int)Props.Image;
                    break;
                case "IPicture":                
                    plPropNum = (int)Props.IPicture;
                    break;                
				default:
					plPropNum = -1;
					break;
			}
		}
		
		public void GetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
		{	//����� 1� (������������) ������ ��� �������� �� ��� ��������������. lPropAlias - ����� ����������
			pbstrPropName = "";
		}

		public void GetPropVal(int lPropNum, ref object pvarPropVal)
		{	//����� 1� ������ �������� ������� 
			pvarPropVal = null;
			switch(lPropNum)
			{
				case (int)Props.ImageSize:
                    pvarPropVal = Convert.ToString(�����������.Width) + 'x' + Convert.ToString(�����������.Height);
					break;
				case (int)Props.Height:
                    pvarPropVal = �����������.Height;
					break;
                case (int)Props.Width:
                    pvarPropVal = �����������.Width;
                    break;
                case (int)Props.BinaryData:
                    {                        
                        Image ReturnImage = (Image)�����������.Clone();
                        TypeConverter imageConverter = TypeDescriptor.GetConverter(ReturnImage.GetType());
                        byte[] imageData = (byte[])imageConverter.ConvertTo(ReturnImage, typeof(byte[]));                        
                        pvarPropVal = Convert.ToBase64String(imageData, Base64FormattingOptions.InsertLineBreaks);
                        ReturnImage.Dispose();                        
                    } 
                    break;
                case (int)Props.Image:
                    pvarPropVal = �����������;
                    break;
                case (int)Props.IPicture:
                    pvarPropVal = GetIPicture();                    
                    break;
			}
            GC.Collect();
            GC.WaitForFullGCComplete();
		}
		
		public void SetPropVal(int lPropNum, ref object varPropVal)
        {	//����� 1� �������� �������� �������
            switch (lPropNum)
            {
                case (int)Props.Image:
                    ����������� =  (Image)((Image)varPropVal).Clone();
                    break;
            }
		}
		
		public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
		{	//����� 1� ������, ����� �������� �������� ��� ������
			pboolPropRead = true; // ��� �������� �������� ��� ������
		}
		
		public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
		{	//����� 1� ������, ����� �������� �������� ��� ������
            switch (lPropNum)
            {
                case (int)Props.Image:
                    pboolPropWrite = true;
                    break;
                default:
                    pboolPropWrite = false;
                    break;
            }            
		}
		#endregion
	
		#region "������"
		enum Methods
		{	//�������� �������������� ������� (�������� ��� �������) ����� ������� ����������
			GetImage = 0, 
			CropImage = 1, 
			ResizeImage = 2, 
			AddWatermark = 3,
            RotateImage = 4,
            SaveImage = 5,
            CloseImage = 6,
            GetImageFromBinararyData = 7,
            LastMethod = 8
		}

		public void GetNMethods(ref int plMethods)
		{	//����� 1� �������� ���������� ��������� �� �� �������
			plMethods = (int)Methods.LastMethod;
		}
		
		public void FindMethod(string bstrMethodName, ref int plMethodNum)
		{	//����� 1� �������� �������� ������������� ������ (��������� ��� �������) �� ����� (��������) ��������� ��� �������
			plMethodNum = -1;
			switch(bstrMethodName)
			{
				case "GetImage":
				case "�������������������":
					plMethodNum = (int)Methods.GetImage;
					break;
				case "CropImage":
				case "�������������������":
					plMethodNum = (int)Methods.CropImage;
					break;
                case "ResizeImage":
                case "�������������������������":
					plMethodNum = (int)Methods.ResizeImage;
					break;
                case "AddWatermark":
				case"�������������������":
					plMethodNum = (int)Methods.AddWatermark;
					break;
                case "ReturnImage":
                case "��������������������":
                    plMethodNum = (int)Methods.RotateImage;
                    break;
                case "SaveImage":
                case "��������������������":
                    plMethodNum = (int)Methods.SaveImage;
                    break;
                case "CloseImage":
                case "�������":
                    plMethodNum = (int)Methods.CloseImage;
                    break;
                case "GetImageFromBinararyData":
                case "�����������������������������������":
                    plMethodNum = (int)Methods.GetImageFromBinararyData;
                    break;

			}
		}
		
		public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
		{	//����� 1� (������������) �������� ��� ������ �� ��� ��������������. lMethodAlias - ����� ��������.
			pbstrMethodName = "";
		}

		public void GetNParams(int lMethodNum, ref int plParams)
		{	//����� 1� �������� ���������� ���������� � ������ (��������� ��� �������)
			switch(lMethodNum)
			{
				case (int)Methods.GetImage:
					plParams = 2;
					break;
				case (int)Methods.CropImage:
					plParams = 5;
					break;
                case (int)Methods.ResizeImage:
					plParams = 5;
					break;
                case (int)Methods.AddWatermark:
					plParams = 3;
					break;
                case (int)Methods.RotateImage:
                    plParams = 3;
                    break;
                case (int)Methods.SaveImage:
                    plParams = 4;
                    break;
                case (int)Methods.CloseImage:
                    plParams = 0;
                    break;
                case (int)Methods.GetImageFromBinararyData:
                    plParams = 1;
                    break;
			}
		}
		
		public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
		{	//����� 1� �������� �������� ���������� ��������� ��� ������� �� ���������
			pvarParamDefValue = null; //��� �������� �� ���������
            switch (lMethodNum)
            {
                case (int)Methods.ResizeImage:
                    switch (lParamNum)
                    {
                        case 4:
                            pvarParamDefValue = 0;
                            break;
                    }                    
                    break;
               	case (int)Methods.GetImage:
                    switch (lParamNum)
                    {
                        case 1:
                            pvarParamDefValue = true;
                            break;
                    }                    
                    break;
            }
		}

		public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
		{	//����� 1� ������, ���������� �� ����� �������� (�.�. �������� ���������� ��� ��������)
			pboolRetValue = true;  //��� ������ � ��� ����� ��������� (�.�. ����� ���������� ��������). 
		}

		public void CallAsProc(int lMethodNum, ref System.Array paParams)
		{	//����� ������� ���������� ��������� ��� ��������. � �������� � ��� ���.
		}

		public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
		{	//����� ������� ���������� ��������� ��� �������.			
			pvarRetValue = 0; //������������ �������� ������ ��� 1�
			
			switch(lMethodNum) //���������� ����� ������
			{
                // ����� "�������������������"
                #region �������������������
                case (int)Methods.GetImage: 
				{                    
                    try
                    {
                        ����������� = null;
                        string ���� = (string)paParams.GetValue(0);                        
                        bool �������� = (bool)paParams.GetValue(1);
                        if (!String.IsNullOrEmpty(����))
                        {                            
                            // �������� ����������� �� �����                                                                                                     
                            byte[] ���������� = File.ReadAllBytes(����);                         
                            MemoryStream ����� = new MemoryStream(����������);                            
                            ����������� = Image.FromStream(�����);                            

                            //����������� = Image.FromFile(����);
                            if (����������� is Metafile)
                            {
                                // ������� �� 2 px � ���� � ������ (�����-�� ���� � MetafilePict)
                                Rectangle ������������������������ = new Rectangle(new Point(0, 0), new Size(�����������.Width - 2, �����������.Height - 2));
                                �������������������(ref ������������������������, ref �����������);
                                ������������������������ = Rectangle.Empty;
                            }
                            ������������ = ����;
                        }
                        else
                        {                        	
                            // �������� ����������� �� ������ ������
                            if (Clipboard.ContainsImage())
                            {                                
                                ����������� = Clipboard.GetImage();                                                                
                            }
                            else if (Clipboard.ContainsData(DataFormats.MetafilePict))
                            {                                
                                if (OpenClipboard(IntPtr.Zero)) // ����� �������� ���� ����������
                                {                                    
                                    if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                                    {                                 
                                        IntPtr ptr = GetClipboardData(CF_ENHMETAFILE);                                     
                                        if (!ptr.Equals(new IntPtr(0)))
                                        {                                            
                                            Metafile metafile = new Metafile(ptr, true);                                         
                                            Image �������������������� = (Image)metafile;                                            
                                                                                        
                                            if (��������)                                            
                                            {
                                            	// ������� �� 2 px � ���� � ������ (�����-�� ���� � MetafilePict)
                                            	Rectangle ������������������������ = new Rectangle(new Point(0, 0), new Size(��������������������.Width - 2, ��������������������.Height - 2));                                            
                                            	�������������������(ref ������������������������, ref ��������������������);
                                            	������������������������ = Rectangle.Empty;
                                            }
                                            else
                                            	����������� = (Image)��������������������.Clone();
                                            
                                            
                                            ��������������������.Dispose();                                            
                                            metafile.Dispose();											
                                        }
                                        ptr = IntPtr.Zero;
                                    }
                                    CloseClipboard();
                                }								
                            }    							
                        }
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }
                    if (����������� == null)
                        ��������("�� ������� �������� �����������");                                        
                    
                    break;
                }
                #endregion
                // ����� ������ "�������������������"					

                // ����� "�����������������������������������"
                #region �����������������������������������
                case (int)Methods.GetImageFromBinararyData: 
				{                    
                    try
                    {
                        ����������� = null;
                        string �������� = (string)paParams.GetValue(0);

                        byte[] imageData = Convert.FromBase64String(��������);
                        
                        msBinary = new System.IO.MemoryStream(imageData);
                        ����������� = System.Drawing.Image.FromStream(msBinary);

                        �������� = null;
                        imageData = null;                        
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }
                    if (����������� == null)
                        ��������("�� ������� �������� ����������� �� �������� ������");                                        
                    break;
                }
                #endregion
                // ����� ����� "�����������������������������������"

                // ����� "�������������������"
                #region �������������������
                case (int)Methods.CropImage:  
				{
                    try
                    {
                        Int32 ���� = Convert.ToInt32(paParams.GetValue(0));
                        Int32 ����� = Convert.ToInt32(paParams.GetValue(1));
                        Int32 ���� = Convert.ToInt32(paParams.GetValue(2));
                        Int32 ��� = Convert.ToInt32(paParams.GetValue(3));
                        bool ���������� = Convert.ToBoolean(paParams.GetValue(4));

                        if (����������)
                        {
                            ���� = �����������.Width / 100 * ����;
                            ����� = �����������.Width / 100 * �����;
                            ���� = �����������.Height / 100 * ����;
                            ��� = �����������.Height / 100 * ���;
                        }

                        Rectangle ������������������������ = new Rectangle();

                        if (���� > 0)
                        {
                            if (���� < �����������.Width)
                            {
                                ������������������������ = new Rectangle(new Point(����, 0), new Size(�����������.Width - ����, �����������.Height));
                                �������������������(ref ������������������������, ref �����������);
                                ������������������������ = Rectangle.Empty;
                            }

                        }
                        if (����� > 0)
                        {
                            if (����� < �����������.Width)
                            {
                                ������������������������ = new Rectangle(new Point(0, 0), new Size(�����������.Width - �����, �����������.Height));
                                �������������������(ref ������������������������, ref �����������);
                                ������������������������ = Rectangle.Empty;
                            }
                        }
                        if (���� > 0)
                        {
                            if (���� < �����������.Height)
                            {
                                ������������������������ = new Rectangle(new Point(0, ����), new Size(�����������.Width, �����������.Height - ����));
                                �������������������(ref ������������������������, ref �����������);
                                ������������������������ = Rectangle.Empty;
                            }
                        }
                        if (��� > 0)
                        {
                            if (��� < �����������.Height)
                            {
                                ������������������������ = new Rectangle(new Point(0, 0), new Size(�����������.Width, �����������.Height - ���));
                                �������������������(ref ������������������������, ref �����������);
                                ������������������������ = Rectangle.Empty;
                            }
                        }
                        
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }
					break;
                }
                #endregion
                // ����� ������ "�������������������"				

                // ����� "�������������������������"
                #region �������������������������
                case (int)Methods.ResizeImage:  //��������� ����� ��� ������ ��������� �� ������
				{
                    try
                    {
                        Int32 ������ = Convert.ToInt32(paParams.GetValue(0));
                        Int32 ������ = Convert.ToInt32(paParams.GetValue(1));
                        Int32 ������� = Convert.ToInt32(paParams.GetValue(2));
                        bool ������������������ = Convert.ToBoolean(paParams.GetValue(3));
                        Int32 �������� = 0;
                        if (paParams.Length > 4)
                            �������� = Convert.ToInt32(paParams.GetValue(4));                        
                        

                        Int32 ����������� = 0;
                        Int32 ����������� = 0;

                        if (������� != 0)
                        {
                            // ��������������� � ���������
                            if (������� > 0)
                            {
                                ����������� = �����������.Width + Convert.ToInt32((((double)�����������.Width / 100) * �������));
                                ����������� = �����������.Height + Convert.ToInt32((((double)�����������.Height / 100) * �������));
                            }
                            else
                            {
                                ����������� = �����������.Width + Convert.ToInt32((((double)�����������.Width / 100) * �������));
                                ����������� = �����������.Height + Convert.ToInt32((((double)�����������.Height / 100) * �������));
                            }                            
                        }
                        else
                        {
                            if (������������������)
                            {
                                Double ����������� = 0;
                                if (������ > 0)
                                    ����������� = (double)������ / (double)�����������.Width;
                                else if (������ > 0)
                                    ����������� = (double)������ / (double)�����������.Height;
                                else
                                    break;
                                
                                ����������� = Convert.ToInt32(�����������.Width * �����������);
                                ����������� = Convert.ToInt32(�����������.Height * �����������);

                            }
                            else
                            {
                                if ((������ == 0) | (������ == 0))
                                    break;

                                ����������� = ������;
                                ����������� = ������;
                            }
                        }

                        ����������� = �������������������������(�����������, �����������, �����������, ��������);
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }

                    break;
                }
                #endregion
                // ����� ������ "�������������������������"

                // ����� "�������������������"
                #region �������������������
                case (int)Methods.AddWatermark: 
				{
                    try
                    {
                        string �������� = Convert.ToString(paParams.GetValue(0));
                        Int32 ������� = Convert.ToInt32(paParams.GetValue(1));
                        Int32 ������������ = Convert.ToInt32(paParams.GetValue(2));

                        // ��������� ������� ����
                        Image ����������� = Image.FromFile(��������);
                        
                        Int32 x = 0; Int32 y = 0;

                        // �������� ���������� �������� ������ ���� �������� �����
                        if (������� == 0)
                        {
                            // ������ �����
                            x = 0;
                            y = 0;
                        }
                        else if (������� == 1)
                        {
                            // ������ �����
                            x = �����������.Width / 2 - �����������.Width / 2;
                            y = 0;
                        }
                        else if (������� == 2)
                        {
                            // ������ ������
                            x = �����������.Width - �����������.Width;
                            y = 0;
                        }
                        else if (������� == 3)
                        {
                            // �����
                            x = Convert.ToInt32(Convert.ToDouble(�����������.Width) / 2 - Convert.ToDouble(�����������.Width) / 2);
                            y = �����������.Height / 2 - �����������.Height / 2;
                        }
                        else if (������� == 4)
                        {
                            // ����� �����
                            x = 0;
                            y = �����������.Height - �����������.Height;
                        }
                        else if (������� == 5)
                        {
                            // ����� �����
                            x = �����������.Width / 2 - �����������.Width / 2;
                            y = �����������.Height - �����������.Height;
                        }
                        else if (������� == 6)
                        {
                            // ����� ������
                            x = �����������.Width - �����������.Width;
                            y = �����������.Height - �����������.Height;
                        }
                        else
                        {
                            // ��� ������ �������� ��������� �� ������
                            x = �����������.Width / 2 - �����������.Width / 2;
                            y = �����������.Height / 2 - �����������.Height / 2;
                        }
                                                                        
                        Graphics g = Graphics.FromImage(�����������);

                        Bitmap transparentWater = new Bitmap(�����������.Width, �����������.Height);
                        Graphics transGraphics = Graphics.FromImage(transparentWater);
                        ColorMatrix tranMatrix = new ColorMatrix();
                        tranMatrix.Matrix33  = ((float)������������/100);  // ������������� ������������

                        ImageAttributes transparentAtt = new ImageAttributes();
                        transparentAtt.SetColorMatrix(tranMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                        transGraphics.DrawImage(�����������, new Rectangle(0, 0, transparentWater.Width, transparentWater.Height), 0, 0, transparentWater.Width, transparentWater.Height, GraphicsUnit.Pixel, transparentAtt);
                        transGraphics.Dispose();                        
                        g.DrawImage(transparentWater, x, y, transparentWater.Width, transparentWater.Height); // ���������� �������� �����                        
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }
					break;
                }
                #endregion
                // ����� ������ "�������������������"

                // ����� "�������������������"
                #region ��������������������
                case (int)Methods.RotateImage:
                {
                    try
                    {
                        bool ��������������������� = (bool)paParams.GetValue(0);
                        bool ������������������� = (bool)paParams.GetValue(1);
                        Int32 ���� = Convert.ToInt32(paParams.GetValue(2));                    
                        if (���������������������)
                            �����������.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        if (�������������������)
                            �����������.RotateFlip(RotateFlipType.RotateNoneFlipY);
                        if (���� == 1)
                            �����������.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        else if (���� == -1)
                            �����������.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);                        
                    }
                    break;
                }
                #endregion
                // ����� ������ "��������������������"

                // ����� "��������������������"
                #region ��������������������
                case (int)Methods.SaveImage: 
                {                    
                    try
                    {
                        string ���� = (string)paParams.GetValue(0);
                        Int32 ������ = Convert.ToInt32(paParams.GetValue(1));
                        Int32 ������ = Convert.ToInt32(paParams.GetValue(2));
                        bool ������������������� = Convert.ToBoolean(paParams.GetValue(3));
                        System.Drawing.Imaging.ImageFormat ����������������� = System.Drawing.Imaging.ImageFormat.Bmp;                        
                        switch (������)
                        {
                            case 1:
                                ����������������� = System.Drawing.Imaging.ImageFormat.Jpeg;
                                break;
                            case 2:
                                ����������������� = System.Drawing.Imaging.ImageFormat.Png;
                                break;
                        }

                        if (������ == 1)
                        {                            
                            if (������ == 0)
                                ������ = 100;
                            EncoderParameter qualityParam = new EncoderParameter(Encoder.Quality, ������);                         
                            // �������� ������ �����
                            ImageCodecInfo[] ������ = ImageCodecInfo.GetImageEncoders();
                            ImageCodecInfo �����JPEG = null;
                            for (int i = 0; i < ������.Length-1;i++)
                            {
                                if (������[i].MimeType == "image/jpeg")
                                {
                                    �����JPEG = ������[i];
                                    break;
                                }
                            }
                            
                            if (�����JPEG == null)
                            {
                                ��������("�� ��������� ����� \"image/jpeg\"");
                                return;
                            }
                            
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = qualityParam;                            
                                                        
                            �����������.Save(����, �����JPEG, encoderParams);                            
                        }
                        else
                        {
                            �����������.Save(����, �����������������);
                        }

                        // ������� �������� ����
                        if (!String.IsNullOrEmpty(������������) && (�������������������))
                        {
                            try
                            {
                                File.Delete(������������);
                            }
                            catch (Exception E)
                            {
                                ��������("�� ������� ������� �������� ����. (" + E.Message + ")");
                            }
                        }
                    }
                    catch (Exception E)
                    {
                        ��������(E.Message);
                    }
                    break;
                }
                #endregion
                // ����� ������ "��������������������"

                // ����� "�������"
                #region �������
                case (int)Methods.CloseImage:
                {
                    msBinary.Dispose();
                    �����������.Dispose();                    
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    break;
                }
                #endregion
                // ����� ������ "�������"
            }
		}
		#endregion

        #region ���������������������
        public void ��������(string ��������������)
        {
            object App1C = null;
            object MessageStatus = null;
            object Status = null;
            //object AppVersion = null;

            object[] ��������� = new object[2];
            ���������[0] = "GraphicsComp: " + ��������������;

            bool ������8X = true;

            try
            {
                App1C = Data1C.Object1C.GetType().InvokeMember("AppDispatch", BindingFlags.GetProperty, null, Data1C.Object1C, null);

                try
                {
                    MessageStatus = Data1C.Object1C.GetType().InvokeMember("MessageStatus", BindingFlags.GetProperty, null, App1C, null);
                    Status = Data1C.Object1C.GetType().InvokeMember("Attention", BindingFlags.GetProperty, null, MessageStatus, null);
                    ���������[1] = Status;
                }
                catch
                {
                    ���������[1] = "!";
                    ������8X = false;
                    Marshal.Release(Marshal.GetIDispatchForObject(Status));
                    Marshal.ReleaseComObject(Status);
                    Status = null;
                    Marshal.Release(Marshal.GetIDispatchForObject(MessageStatus));
                    Marshal.ReleaseComObject(MessageStatus);
                    MessageStatus = null;
                }

                Data1C.Object1C.GetType().InvokeMember("Message", BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static, null, App1C, ���������);
            }
            finally
            {
                if (������8X)
                {
                    Marshal.Release(Marshal.GetIDispatchForObject(Status));
                    Marshal.ReleaseComObject(Status);
                    Status = null;
                    Marshal.Release(Marshal.GetIDispatchForObject(MessageStatus));
                    Marshal.ReleaseComObject(MessageStatus);
                    MessageStatus = null;
                }
                Marshal.Release(Marshal.GetIDispatchForObject(App1C));
                Marshal.ReleaseComObject(App1C);
                App1C = null;
            }
            throw new COMException();      
        }

        public void �������������������(ref Rectangle �������, ref Image �������������������)
        {            
        	using (Bitmap ������������������ = new Bitmap(�������������������))
        	{
        		using (Bitmap ���������������� = ������������������.Clone(�������, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
        		{	
            		����������� = Image.FromHbitmap(����������������.GetHbitmap());            		
        		}
        	}
        }
        
        public Image �������������������������(Int32 �����������, Int32 �����������, Image �������������������, Int32 ��������)
        {
        
            Bitmap b = new Bitmap(�����������, �����������);            
            Graphics g = Graphics.FromImage((Image)b);
            if (�������� == 0)
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.SmoothingMode = SmoothingMode.HighSpeed;                
            }            
            g.DrawImage(�������������������, 0, 0, �����������, �����������);
            g.Dispose();
            
            return (Image)b;
        }

        public object GetIPicture()
        {
            try
            {
                ImageConverter ConvertImg = new ImageConverter();
                object result = ConvertImg.ImageToIpicture(�����������);
                ConvertImg.Dispose();
                return result;
            }
            catch (Exception E)
            {
                ��������("�� ������� �������� ������ \"IPicture\":" + E.Message);
                return null;
            }
        }

        #endregion
    }
}
