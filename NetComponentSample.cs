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
            Изображение.Dispose();
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
					  
		#region "Переменные"		
        public Image Изображение;
        public object App1C = null;
        public object Pic1C = null;
        public object ДвоичныеДанные = null;
        private MemoryStream msBinary = null;        
        string ИсходныйФайл;
		#endregion
															   
		public void RegisterExtensionAs(ref string bstrExtensionName)
		{            
			bstrExtensionName = c_AddinName;            
		}
		
		#region "Свойства"
		enum Props
		{   //Числовые идентификаторы свойств нашей внешней компоненты
			ImageSize = 0,  
            Width = 1, 
            Height = 2,
            BinaryData = 3,
            Image = 4,
            IPicture = 5,
            LastProp = 6
		}
         
		public void GetNProps(ref int plProps)
		{	//Здесь 1С получает количество доступных из ВК свойств
			plProps = (int)Props.LastProp;
		}
        
		public void FindProp(string bstrPropName, ref int plPropNum)
		{	//Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
			switch(bstrPropName)
			{
				case "ImageSize": 
				case "РазмерИзображения":
					plPropNum = (int)Props.ImageSize;
					break;
                case "Height":
				case "Высота":
					plPropNum = (int)Props.Height;
					break;
                case "Width":
                case "Ширина":
                    plPropNum = (int)Props.Width;
                    break;
                case "BinaryData":
                case "ДвоичныеДанные":
                    plPropNum = (int)Props.BinaryData;
                    break;
                case "Image":
                case "Изображение":
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
		{	//Здесь 1С (теоретически) узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
			pbstrPropName = "";
		}

		public void GetPropVal(int lPropNum, ref object pvarPropVal)
		{	//Здесь 1С узнает значения свойств 
			pvarPropVal = null;
			switch(lPropNum)
			{
				case (int)Props.ImageSize:
                    pvarPropVal = Convert.ToString(Изображение.Width) + 'x' + Convert.ToString(Изображение.Height);
					break;
				case (int)Props.Height:
                    pvarPropVal = Изображение.Height;
					break;
                case (int)Props.Width:
                    pvarPropVal = Изображение.Width;
                    break;
                case (int)Props.BinaryData:
                    {                        
                        Image ReturnImage = (Image)Изображение.Clone();
                        TypeConverter imageConverter = TypeDescriptor.GetConverter(ReturnImage.GetType());
                        byte[] imageData = (byte[])imageConverter.ConvertTo(ReturnImage, typeof(byte[]));                        
                        pvarPropVal = Convert.ToBase64String(imageData, Base64FormattingOptions.InsertLineBreaks);
                        ReturnImage.Dispose();                        
                    } 
                    break;
                case (int)Props.Image:
                    pvarPropVal = Изображение;
                    break;
                case (int)Props.IPicture:
                    pvarPropVal = GetIPicture();                    
                    break;
			}
            GC.Collect();
            GC.WaitForFullGCComplete();
		}
		
		public void SetPropVal(int lPropNum, ref object varPropVal)
        {	//Здесь 1С изменяет значения свойств
            switch (lPropNum)
            {
                case (int)Props.Image:
                    Изображение =  (Image)((Image)varPropVal).Clone();
                    break;
            }
		}
		
		public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
		{	//Здесь 1С узнает, какие свойства доступны для чтения
			pboolPropRead = true; // Все свойства доступны для чтения
		}
		
		public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
		{	//Здесь 1С узнает, какие свойства доступны для записи
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
	
		#region "Методы"
		enum Methods
		{	//Числовые идентификаторы методов (процедур или функций) нашей внешней компоненты
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
		{	//Здесь 1С получает количество доступных из ВК методов
			plMethods = (int)Methods.LastMethod;
		}
		
		public void FindMethod(string bstrMethodName, ref int plMethodNum)
		{	//Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
			plMethodNum = -1;
			switch(bstrMethodName)
			{
				case "GetImage":
				case "ПолучитьИзображение":
					plMethodNum = (int)Methods.GetImage;
					break;
				case "CropImage":
				case "ОбрезатьИзображение":
					plMethodNum = (int)Methods.CropImage;
					break;
                case "ResizeImage":
                case "ИзменитьРазмерИзображения":
					plMethodNum = (int)Methods.ResizeImage;
					break;
                case "AddWatermark":
				case"ДобавитьВодянойЗнак":
					plMethodNum = (int)Methods.AddWatermark;
					break;
                case "ReturnImage":
                case "ПовернутьИзображение":
                    plMethodNum = (int)Methods.RotateImage;
                    break;
                case "SaveImage":
                case "СохранитьИзображение":
                    plMethodNum = (int)Methods.SaveImage;
                    break;
                case "CloseImage":
                case "Закрыть":
                    plMethodNum = (int)Methods.CloseImage;
                    break;
                case "GetImageFromBinararyData":
                case "ПолучитьИзображениеИзДвоичныхДанных":
                    plMethodNum = (int)Methods.GetImageFromBinararyData;
                    break;

			}
		}
		
		public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
		{	//Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
			pbstrMethodName = "";
		}

		public void GetNParams(int lMethodNum, ref int plParams)
		{	//Здесь 1С получает количество параметров у метода (процедуры или функции)
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
		{	//Здесь 1С получает значения параметров процедуры или функции по умолчанию
			pvarParamDefValue = null; //Нет значений по умолчанию
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
		{	//Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
			pboolRetValue = true;  //Все методы у нас будут функциями (т.е. будут возвращать значение). 
		}

		public void CallAsProc(int lMethodNum, ref System.Array paParams)
		{	//Здесь внешняя компонента выполняет код процедур. А процедур у нас нет.
		}

		public void CallAsFunc(int lMethodNum, ref object pvarRetValue, ref System.Array paParams)
		{	//Здесь внешняя компонента выполняет код функций.			
			pvarRetValue = 0; //Возвращаемое значение метода для 1С
			
			switch(lMethodNum) //Порядковый номер метода
			{
                // Метод "ПолучитьИзображение"
                #region ПолучитьИзображение
                case (int)Methods.GetImage: 
				{                    
                    try
                    {
                        Изображение = null;
                        string Файл = (string)paParams.GetValue(0);                        
                        bool Обрезать = (bool)paParams.GetValue(1);
                        if (!String.IsNullOrEmpty(Файл))
                        {                            
                            // Получаем изображение из файла                                                                                                     
                            byte[] Содержимое = File.ReadAllBytes(Файл);                         
                            MemoryStream Поток = new MemoryStream(Содержимое);                            
                            Изображение = Image.FromStream(Поток);                            

                            //Изображение = Image.FromFile(Файл);
                            if (Изображение is Metafile)
                            {
                                // Обрежем по 2 px с низу и справа (какой-то глюк у MetafilePict)
                                Rectangle ОбластьНовогоИзображения = new Rectangle(new Point(0, 0), new Size(Изображение.Width - 2, Изображение.Height - 2));
                                ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref Изображение);
                                ОбластьНовогоИзображения = Rectangle.Empty;
                            }
                            ИсходныйФайл = Файл;
                        }
                        else
                        {                        	
                            // Получаем изображение из буфера обмена
                            if (Clipboard.ContainsImage())
                            {                                
                                Изображение = Clipboard.GetImage();                                                                
                            }
                            else if (Clipboard.ContainsData(DataFormats.MetafilePict))
                            {                                
                                if (OpenClipboard(IntPtr.Zero)) // Хэндл главного окна приложения
                                {                                    
                                    if (IsClipboardFormatAvailable(CF_ENHMETAFILE))
                                    {                                 
                                        IntPtr ptr = GetClipboardData(CF_ENHMETAFILE);                                     
                                        if (!ptr.Equals(new IntPtr(0)))
                                        {                                            
                                            Metafile metafile = new Metafile(ptr, true);                                         
                                            Image ВременноеИзображение = (Image)metafile;                                            
                                                                                        
                                            if (Обрезать)                                            
                                            {
                                            	// Обрежем по 2 px с низу и справа (какой-то глюк у MetafilePict)
                                            	Rectangle ОбластьНовогоИзображения = new Rectangle(new Point(0, 0), new Size(ВременноеИзображение.Width - 2, ВременноеИзображение.Height - 2));                                            
                                            	ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref ВременноеИзображение);
                                            	ОбластьНовогоИзображения = Rectangle.Empty;
                                            }
                                            else
                                            	Изображение = (Image)ВременноеИзображение.Clone();
                                            
                                            
                                            ВременноеИзображение.Dispose();                                            
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
                        Сообщить(E.Message);
                    }
                    if (Изображение == null)
                        Сообщить("Не удалось получить изображение");                                        
                    
                    break;
                }
                #endregion
                // конец метода "ПолучитьИзображение"					

                // Метод "ПолучитьИзображениеИзДвоичныхДанных"
                #region ПолучитьИзображениеИзДвоичныхДанных
                case (int)Methods.GetImageFromBinararyData: 
				{                    
                    try
                    {
                        Изображение = null;
                        string ДвДанные = (string)paParams.GetValue(0);

                        byte[] imageData = Convert.FromBase64String(ДвДанные);
                        
                        msBinary = new System.IO.MemoryStream(imageData);
                        Изображение = System.Drawing.Image.FromStream(msBinary);

                        ДвДанные = null;
                        imageData = null;                        
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);
                    }
                    if (Изображение == null)
                        Сообщить("Не удалось получить изображение из двоичных данных");                                        
                    break;
                }
                #endregion
                // Конец метод "ПолучитьИзображениеИзДвоичныхДанных"

                // Метод "ОбрезатьИзображение"
                #region ОбрезатьИзображение
                case (int)Methods.CropImage:  
				{
                    try
                    {
                        Int32 Лево = Convert.ToInt32(paParams.GetValue(0));
                        Int32 Право = Convert.ToInt32(paParams.GetValue(1));
                        Int32 Верх = Convert.ToInt32(paParams.GetValue(2));
                        Int32 Низ = Convert.ToInt32(paParams.GetValue(3));
                        bool ВПроцентах = Convert.ToBoolean(paParams.GetValue(4));

                        if (ВПроцентах)
                        {
                            Лево = Изображение.Width / 100 * Лево;
                            Право = Изображение.Width / 100 * Право;
                            Верх = Изображение.Height / 100 * Верх;
                            Низ = Изображение.Height / 100 * Низ;
                        }

                        Rectangle ОбластьНовогоИзображения = new Rectangle();

                        if (Лево > 0)
                        {
                            if (Лево < Изображение.Width)
                            {
                                ОбластьНовогоИзображения = new Rectangle(new Point(Лево, 0), new Size(Изображение.Width - Лево, Изображение.Height));
                                ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref Изображение);
                                ОбластьНовогоИзображения = Rectangle.Empty;
                            }

                        }
                        if (Право > 0)
                        {
                            if (Право < Изображение.Width)
                            {
                                ОбластьНовогоИзображения = new Rectangle(new Point(0, 0), new Size(Изображение.Width - Право, Изображение.Height));
                                ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref Изображение);
                                ОбластьНовогоИзображения = Rectangle.Empty;
                            }
                        }
                        if (Верх > 0)
                        {
                            if (Верх < Изображение.Height)
                            {
                                ОбластьНовогоИзображения = new Rectangle(new Point(0, Верх), new Size(Изображение.Width, Изображение.Height - Верх));
                                ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref Изображение);
                                ОбластьНовогоИзображения = Rectangle.Empty;
                            }
                        }
                        if (Низ > 0)
                        {
                            if (Низ < Изображение.Height)
                            {
                                ОбластьНовогоИзображения = new Rectangle(new Point(0, 0), new Size(Изображение.Width, Изображение.Height - Низ));
                                ОбрезатьИзображение(ref ОбластьНовогоИзображения, ref Изображение);
                                ОбластьНовогоИзображения = Rectangle.Empty;
                            }
                        }
                        
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);
                    }
					break;
                }
                #endregion
                // конец метода "ОбрезатьИзображение"				

                // Метод "ИзменитьРазмерИзображения"
                #region ИзменитьРазмерИзображения
                case (int)Methods.ResizeImage:  //Реализуем метод для показа сообщения об ошибке
				{
                    try
                    {
                        Int32 Ширина = Convert.ToInt32(paParams.GetValue(0));
                        Int32 Высота = Convert.ToInt32(paParams.GetValue(1));
                        Int32 Процент = Convert.ToInt32(paParams.GetValue(2));
                        bool СохранятьПропорции = Convert.ToBoolean(paParams.GetValue(3));
                        Int32 Качество = 0;
                        if (paParams.Length > 4)
                            Качество = Convert.ToInt32(paParams.GetValue(4));                        
                        

                        Int32 НоваяШирина = 0;
                        Int32 НоваяВысота = 0;

                        if (Процент != 0)
                        {
                            // Масштабирование в процентах
                            if (Процент > 0)
                            {
                                НоваяШирина = Изображение.Width + Convert.ToInt32((((double)Изображение.Width / 100) * Процент));
                                НоваяВысота = Изображение.Height + Convert.ToInt32((((double)Изображение.Height / 100) * Процент));
                            }
                            else
                            {
                                НоваяШирина = Изображение.Width + Convert.ToInt32((((double)Изображение.Width / 100) * Процент));
                                НоваяВысота = Изображение.Height + Convert.ToInt32((((double)Изображение.Height / 100) * Процент));
                            }                            
                        }
                        else
                        {
                            if (СохранятьПропорции)
                            {
                                Double Коэффициент = 0;
                                if (Ширина > 0)
                                    Коэффициент = (double)Ширина / (double)Изображение.Width;
                                else if (Высота > 0)
                                    Коэффициент = (double)Высота / (double)Изображение.Height;
                                else
                                    break;
                                
                                НоваяШирина = Convert.ToInt32(Изображение.Width * Коэффициент);
                                НоваяВысота = Convert.ToInt32(Изображение.Height * Коэффициент);

                            }
                            else
                            {
                                if ((Ширина == 0) | (Высота == 0))
                                    break;

                                НоваяШирина = Ширина;
                                НоваяВысота = Высота;
                            }
                        }

                        Изображение = ИзменитьРазмерИзображения(НоваяШирина, НоваяВысота, Изображение, Качество);
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);
                    }

                    break;
                }
                #endregion
                // конец метода "ИзменитьРазмерИзображения"

                // Метод "ДобавитьВодянойЗнак"
                #region ДобавитьВодянойЗнак
                case (int)Methods.AddWatermark: 
				{
                    try
                    {
                        string ИмяФайла = Convert.ToString(paParams.GetValue(0));
                        Int32 Позиция = Convert.ToInt32(paParams.GetValue(1));
                        Int32 Прозрачность = Convert.ToInt32(paParams.GetValue(2));

                        // Открываем водяной знак
                        Image ВодянойЗнак = Image.FromFile(ИмяФайла);
                        
                        Int32 x = 0; Int32 y = 0;

                        // Получаем координаты верхнего левого угла водяного знака
                        if (Позиция == 0)
                        {
                            // Сверху слева
                            x = 0;
                            y = 0;
                        }
                        else if (Позиция == 1)
                        {
                            // Сверху центр
                            x = Изображение.Width / 2 - ВодянойЗнак.Width / 2;
                            y = 0;
                        }
                        else if (Позиция == 2)
                        {
                            // Сверху справа
                            x = Изображение.Width - ВодянойЗнак.Width;
                            y = 0;
                        }
                        else if (Позиция == 3)
                        {
                            // Центр
                            x = Convert.ToInt32(Convert.ToDouble(Изображение.Width) / 2 - Convert.ToDouble(ВодянойЗнак.Width) / 2);
                            y = Изображение.Height / 2 - ВодянойЗнак.Height / 2;
                        }
                        else if (Позиция == 4)
                        {
                            // Снизу слева
                            x = 0;
                            y = Изображение.Height - ВодянойЗнак.Height;
                        }
                        else if (Позиция == 5)
                        {
                            // Снизу центр
                            x = Изображение.Width / 2 - ВодянойЗнак.Width / 2;
                            y = Изображение.Height - ВодянойЗнак.Height;
                        }
                        else if (Позиция == 6)
                        {
                            // Снизу справа
                            x = Изображение.Width - ВодянойЗнак.Width;
                            y = Изображение.Height - ВодянойЗнак.Height;
                        }
                        else
                        {
                            // При другом значении размещаем по центру
                            x = Изображение.Width / 2 - ВодянойЗнак.Width / 2;
                            y = Изображение.Height / 2 - ВодянойЗнак.Height / 2;
                        }
                                                                        
                        Graphics g = Graphics.FromImage(Изображение);

                        Bitmap transparentWater = new Bitmap(ВодянойЗнак.Width, ВодянойЗнак.Height);
                        Graphics transGraphics = Graphics.FromImage(transparentWater);
                        ColorMatrix tranMatrix = new ColorMatrix();
                        tranMatrix.Matrix33  = ((float)Прозрачность/100);  // устанавливаем прозрачность

                        ImageAttributes transparentAtt = new ImageAttributes();
                        transparentAtt.SetColorMatrix(tranMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
                        transGraphics.DrawImage(ВодянойЗнак, new Rectangle(0, 0, transparentWater.Width, transparentWater.Height), 0, 0, transparentWater.Width, transparentWater.Height, GraphicsUnit.Pixel, transparentAtt);
                        transGraphics.Dispose();                        
                        g.DrawImage(transparentWater, x, y, transparentWater.Width, transparentWater.Height); // размещение водяного знака                        
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);
                    }
					break;
                }
                #endregion
                // конец метода "ДобавитьВодянойЗнак"

                // Метод "ПовернутИзображение"
                #region ПовернутьИзображение
                case (int)Methods.RotateImage:
                {
                    try
                    {
                        bool ОтразитьПоГоризонтали = (bool)paParams.GetValue(0);
                        bool ОтразитьПоВертикали = (bool)paParams.GetValue(1);
                        Int32 Угол = Convert.ToInt32(paParams.GetValue(2));                    
                        if (ОтразитьПоГоризонтали)
                            Изображение.RotateFlip(RotateFlipType.RotateNoneFlipX);
                        if (ОтразитьПоВертикали)
                            Изображение.RotateFlip(RotateFlipType.RotateNoneFlipY);
                        if (Угол == 1)
                            Изображение.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        else if (Угол == -1)
                            Изображение.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);                        
                    }
                    break;
                }
                #endregion
                // конец метода "ПовернутьИзображение"

                // Метод "СохранитьИзображение"
                #region СохранитьИзображение
                case (int)Methods.SaveImage: 
                {                    
                    try
                    {
                        string Путь = (string)paParams.GetValue(0);
                        Int32 Формат = Convert.ToInt32(paParams.GetValue(1));
                        Int32 Сжатие = Convert.ToInt32(paParams.GetValue(2));
                        bool УдалятьИсходныйФайл = Convert.ToBoolean(paParams.GetValue(3));
                        System.Drawing.Imaging.ImageFormat ФорматИзображения = System.Drawing.Imaging.ImageFormat.Bmp;                        
                        switch (Формат)
                        {
                            case 1:
                                ФорматИзображения = System.Drawing.Imaging.ImageFormat.Jpeg;
                                break;
                            case 2:
                                ФорматИзображения = System.Drawing.Imaging.ImageFormat.Png;
                                break;
                        }

                        if (Формат == 1)
                        {                            
                            if (Сжатие == 0)
                                Сжатие = 100;
                            EncoderParameter qualityParam = new EncoderParameter(Encoder.Quality, Сжатие);                         
                            // Получаем нужный кодек
                            ImageCodecInfo[] Кодеки = ImageCodecInfo.GetImageEncoders();
                            ImageCodecInfo КодекJPEG = null;
                            for (int i = 0; i < Кодеки.Length-1;i++)
                            {
                                if (Кодеки[i].MimeType == "image/jpeg")
                                {
                                    КодекJPEG = Кодеки[i];
                                    break;
                                }
                            }
                            
                            if (КодекJPEG == null)
                            {
                                Сообщить("Не обнаружен кодек \"image/jpeg\"");
                                return;
                            }
                            
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = qualityParam;                            
                                                        
                            Изображение.Save(Путь, КодекJPEG, encoderParams);                            
                        }
                        else
                        {
                            Изображение.Save(Путь, ФорматИзображения);
                        }

                        // Удаляем исходный файл
                        if (!String.IsNullOrEmpty(ИсходныйФайл) && (УдалятьИсходныйФайл))
                        {
                            try
                            {
                                File.Delete(ИсходныйФайл);
                            }
                            catch (Exception E)
                            {
                                Сообщить("Не удалось удалить исходный файл. (" + E.Message + ")");
                            }
                        }
                    }
                    catch (Exception E)
                    {
                        Сообщить(E.Message);
                    }
                    break;
                }
                #endregion
                // конец метода "СохранитьИзображение"

                // Метод "Закрыть"
                #region Закрыть
                case (int)Methods.CloseImage:
                {
                    msBinary.Dispose();
                    Изображение.Dispose();                    
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    break;
                }
                #endregion
                // конец метода "Закрыть"
            }
		}
		#endregion

        #region ДополнительныеФункции
        public void Сообщить(string ТекстСообщения)
        {
            object App1C = null;
            object MessageStatus = null;
            object Status = null;
            //object AppVersion = null;

            object[] Параметры = new object[2];
            Параметры[0] = "GraphicsComp: " + ТекстСообщения;

            bool Версия8X = true;

            try
            {
                App1C = Data1C.Object1C.GetType().InvokeMember("AppDispatch", BindingFlags.GetProperty, null, Data1C.Object1C, null);

                try
                {
                    MessageStatus = Data1C.Object1C.GetType().InvokeMember("MessageStatus", BindingFlags.GetProperty, null, App1C, null);
                    Status = Data1C.Object1C.GetType().InvokeMember("Attention", BindingFlags.GetProperty, null, MessageStatus, null);
                    Параметры[1] = Status;
                }
                catch
                {
                    Параметры[1] = "!";
                    Версия8X = false;
                    Marshal.Release(Marshal.GetIDispatchForObject(Status));
                    Marshal.ReleaseComObject(Status);
                    Status = null;
                    Marshal.Release(Marshal.GetIDispatchForObject(MessageStatus));
                    Marshal.ReleaseComObject(MessageStatus);
                    MessageStatus = null;
                }

                Data1C.Object1C.GetType().InvokeMember("Message", BindingFlags.Public | BindingFlags.InvokeMethod | BindingFlags.Static, null, App1C, Параметры);
            }
            finally
            {
                if (Версия8X)
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

        public void ОбрезатьИзображение(ref Rectangle Область, ref Image ИсходноеИзображение)
        {            
        	using (Bitmap ОригинальныйБитмап = new Bitmap(ИсходноеИзображение))
        	{
        		using (Bitmap ОбрезанныйБитмап = ОригинальныйБитмап.Clone(Область, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
        		{	
            		Изображение = Image.FromHbitmap(ОбрезанныйБитмап.GetHbitmap());            		
        		}
        	}
        }
        
        public Image ИзменитьРазмерИзображения(Int32 НоваяШирина, Int32 НоваяВысота, Image ИсходноеИзображение, Int32 Качество)
        {
        
            Bitmap b = new Bitmap(НоваяШирина, НоваяВысота);            
            Graphics g = Graphics.FromImage((Image)b);
            if (Качество == 0)
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.CompositingQuality = CompositingQuality.HighSpeed;
                g.SmoothingMode = SmoothingMode.HighSpeed;                
            }            
            g.DrawImage(ИсходноеИзображение, 0, 0, НоваяШирина, НоваяВысота);
            g.Dispose();
            
            return (Image)b;
        }

        public object GetIPicture()
        {
            try
            {
                ImageConverter ConvertImg = new ImageConverter();
                object result = ConvertImg.ImageToIpicture(Изображение);
                ConvertImg.Dispose();
                return result;
            }
            catch (Exception E)
            {
                Сообщить("Не удалось получить объект \"IPicture\":" + E.Message);
                return null;
            }
        }

        #endregion
    }
}
