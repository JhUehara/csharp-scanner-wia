using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WIA;
using System.Windows.Forms;

namespace ScannerDemo
{
    class Scanner
    {
        private readonly DeviceInfo _deviceInfo;
        private int resolution = 150;
        private int width_pixel = 1250;
        private int height_pixel = 1700;
        private int color_mode = 1;

        public Scanner(DeviceInfo deviceInfo)
        {
            this._deviceInfo = deviceInfo;
        }
 
        /// <summary>
        /// Scanear no formato PNG
        /// </summary>
        /// <returns></returns>
        public ImageFile ScanPNG()
        {
            // Conecte-se ao dispositivo e instrua-o a digitalizar
            // Conecte-se ao dispositivo
            var device = this._deviceInfo.Connect();

            // Selecione o Scanner
            CommonDialogClass dlg = new CommonDialogClass(); 

            var item = device.Items[1];

            try
            {
                AdjustScannerSettings(item, resolution, 0, 0, width_pixel, height_pixel, 0, 0, color_mode);

                object scanResult = dlg.ShowTransfer(item, WIA.FormatID.wiaFormatPNG, true);

                if(scanResult != null)
                {
                    var imageFile = (ImageFile)scanResult;

                    // Retorna o imageFile
                    return imageFile;
                }
            }
            catch (COMException e)
            {
                // Exibe a mensagem de erro no console
                Console.WriteLine(e.ToString());

                uint errorCode = (uint)e.ErrorCode;

                // 2 das exceções mais comuns
                if (errorCode ==  0x80210006)
                {
                    MessageBox.Show("O scanner está ocupado ou indisponível");
                }else if(errorCode == 0x80210064)
                {
                    MessageBox.Show("O processo de digitalização foi cancelado.");
                }else
                {
                    MessageBox.Show("Um erro não detectado ocorreu, verifique no console","Erro",MessageBoxButtons.OK);
                }
            }

            return new ImageFile();
        }

        /// <summary>
        /// Scanear no formato JPEG
        /// </summary>
        /// <returns></returns>
        public ImageFile ScanJPEG()
        {
            // Conecte-se ao dispositivo e instrua-o a digitalizar
            // Conecte-se ao dispositivo
            var device = this._deviceInfo.Connect();

            // Selecione o Scanner
            CommonDialogClass dlg = new CommonDialogClass();

            var item = device.Items[1];

            try
            {
                AdjustScannerSettings(item, resolution, 0, 0, width_pixel, height_pixel, 0, 0, color_mode);

                object scanResult = dlg.ShowTransfer(item, WIA.FormatID.wiaFormatJPEG, true);

                if (scanResult != null)
                {
                    var imageFile = (ImageFile)scanResult;

                    // Retorna o imageFile
                    return imageFile;
                }
            }
            catch (COMException e)
            {
                // Exibe a mensagem de erro no console
                Console.WriteLine(e.ToString());

                uint errorCode = (uint)e.ErrorCode;

                // 2 das exceções mais comuns
                if (errorCode == 0x80210006)
                {
                    MessageBox.Show("O scanner está ocupado ou indisponível");
                }
                else if (errorCode == 0x80210064)
                {
                    MessageBox.Show("O processo de digitalização foi cancelado.");
                }
                else
                {
                    MessageBox.Show("Um erro não detectado ocorreu, verifique no console", "Erro", MessageBoxButtons.OK);
                }
            }

            return new ImageFile();
        }

        /// <summary>
        /// Scannear no formato TIFF
        /// </summary>
        /// <returns></returns>
        public ImageFile ScanTIFF()
        {
            // Conecte-se ao dispositivo e instrua-o a digitalizar
            // Conecte-se ao dispositivo
            var device = this._deviceInfo.Connect();

            // Selecione o Scanner
            CommonDialogClass dlg = new CommonDialogClass();

            var item = device.Items[1];

            try
            {
                AdjustScannerSettings(item, resolution, 0, 0, width_pixel, height_pixel, 0, 0, color_mode);

                object scanResult = dlg.ShowTransfer(item, WIA.FormatID.wiaFormatTIFF, true);

                if (scanResult != null)
                {
                    var imageFile = (ImageFile)scanResult;

                    // Retorna o imageFile
                    return imageFile;
                }
            }
            catch (COMException e)
            {
                // Exibe a mensagem de erro no console
                Console.WriteLine(e.ToString());

                uint errorCode = (uint)e.ErrorCode;

                // 2 das exceções mais comuns
                if (errorCode == 0x80210006)
                {
                    MessageBox.Show("O scanner está ocupado ou indisponível");
                }
                else if (errorCode == 0x80210064)
                {
                    MessageBox.Show("O processo de digitalização foi cancelado.");
                }
                else
                {
                    MessageBox.Show("Um erro não detectado ocorreu, verifique no console", "Erro", MessageBoxButtons.OK);
                }
            }

            return new ImageFile();
        }

        /// <summary>
        ///  Ajusta as configurações do scanner com os parâmetros fornecidos.
        /// </summary>
        /// <param name="scannnerItem">Expects a </param>
        /// <param name="scanResolutionDPI">Forneça a resolução de DPI que deve ser usada, por exemplo, 150</param>
        /// <param name="scanStartLeftPixel"></param>
        /// <param name="scanStartTopPixel"></param>
        /// <param name="scanWidthPixels"></param>
        /// <param name="scanHeightPixels"></param>
        /// <param name="brightnessPercents"></param>
        /// <param name="contrastPercents">Modifique o percentual de contraste</param>
        /// <param name="colorMode">Definir o modo de cor</param>
        private static void AdjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanStartLeftPixel, int scanStartTopPixel, int scanWidthPixels, int scanHeightPixels, int brightnessPercents, int contrastPercents, int colorMode)
        {
            const string WIA_SCAN_COLOR_MODE = "6146";
            const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
            const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
            const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
            const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
            const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
            const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
            const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
            const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, scanStartLeftPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, scanStartTopPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidthPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeightPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, brightnessPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrastPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_COLOR_MODE, colorMode);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="propName"></param>
        /// <param name="propValue"></param>
        private static void SetWIAProperty(IProperties properties, object propName, object propValue)
        {
            Property prop = properties.get_Item(ref propName);
            prop.set_Value(ref propValue);
        }

        /// <summary>
        /// Declarar o método ToString
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return (string) this._deviceInfo.Properties["Name"].get_Value();
        }
         
    }
}
