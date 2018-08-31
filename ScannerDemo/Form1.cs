using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WIA;

namespace ScannerDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ListScanners();

            // Define o diretorio padrão para salvar a digitalização
            textBox1.Text = Directory.GetCurrentDirectory();
            // Define JPEG como default
            comboBox1.SelectedIndex = 1;
            // Define RG como default
            cbDoc.SelectedIndex = 0;

        }

        private void ListScanners()
        {
            // Limpa a ListBox.
            listBox1.Items.Clear();

            // Cria um instancia do scanner
            var deviceManager = new DeviceManager();

            // Loop para adicionar os scanners na listbox
            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                // adiciona somente se for scanner
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                } 
                // Adicione o scanner à listbox
                // Importante: nós armazenamos um objeto do tipo scanner (o método ToString retorna o nome do scanner)
                listBox1.Items.Add(
                    new Scanner(deviceManager.DeviceInfos[i])
                );
            }
        }

        // Digitaliza o documento, e altera o nome do arquivo para salvar de acordo com as seleções
        private void button1_Click(object sender, EventArgs e)
        {
            string digDoc = mtxtCpf.Text.ToString();
            digDoc += "_";
            digDoc += cbDoc.Text;
            textBox2.Text = digDoc;
            Task.Factory.StartNew(StartScanning).ContinueWith(result => TriggerScan());
        }

        private void TriggerScan()
        {
            Console.WriteLine("Imagem digitalizada com sucesso!");
        }

        public void StartScanning()
        {
            Scanner device = null;

            this.Invoke(new MethodInvoker(delegate ()
            {
                device = listBox1.SelectedItem as Scanner;
            }));

            if (device == null)
            {
                MessageBox.Show("Primeiro, você precisa selecionar um scanner na lista",
                                "Atenção",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }else if(String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Nomeie o arquivo",
                                "Atenção",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ImageFile image = new ImageFile();
            string imageExtension = "";

            this.Invoke(new MethodInvoker(delegate ()
            {
                switch (comboBox1.SelectedIndex)
                {
                    case 0:
                        image = device.ScanPNG();
                        imageExtension = ".png";
                        break;
                    case 1:
                        image = device.ScanJPEG();
                        imageExtension = ".jpeg";
                        break;
                    case 2:
                        image = device.ScanTIFF();
                        imageExtension = ".tiff";
                        break;
                }
            }));
            
            
            // salve a imagem
            var path = Path.Combine(textBox1.Text, textBox2.Text, mtxtCpf.Text + imageExtension);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            if (File.Exists(path))
            {
                for (int i = 1; File.Exists(path); i++)
                {
                    image.SaveFile(path + i);
                }
            }
            image.SaveFile(path);

            pictureBox1.Image = new Bitmap(path);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            DialogResult result = folderDlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                textBox1.Text = folderDlg.SelectedPath;
            }
        }
         
    }
}
