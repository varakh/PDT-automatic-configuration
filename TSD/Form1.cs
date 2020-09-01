using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Media;
using System.Windows.Forms;
using Microsoft.Win32;
using Symbol.Fusion;
using Symbol.Fusion.WLAN;
using Symbol.Exceptions;
using Terranova.API;

namespace TSD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //установка значения параметра AutoEnter ключа реестра 
            //(необходимо для работы scanwedge, чтобы после считывания штрих-кода происходило автонажатие клавиши Enter) 
            registryKeySet();

            //запуск scanwedge
            scanwedgeStart();

            //проверка wifi-адаптера и его настроек
            checkWifiAdapter();

            //создание rdp-файла
            rdpFileCreate();

            //вывод статуса и выход из программы
            showStatus();
        }

        
        private Config myConfig = null;
        
        private WLAN myWlan = null;

        //диалог для запроса от пользователя номера ТСД
        private static string ShowDialog(string caption, string text)
        {
            Form prompt = new Form();
            prompt.Width = 280;
            prompt.Height = 150;
            prompt.Text = caption;
            Label textLabel = new Label() { Left = 16, Top = 20, Width = 240, Text = text };
            TextBox textBox = new TextBox() { Left = 16, Top = 40, Width = 240, TabIndex = 0, TabStop = true };
            Button confirmation = new Button() { Text = "ОК", Left = 16, Width = 80, Top = 72, TabIndex = 1, TabStop = true };
            confirmation.Click += (sender, e) =>
            {   //если введённый текст равен нулю или содержит знаки, отличные от нуля, запрос повторяется
                if ((textBox.Text == "0") | (!System.Text.RegularExpressions.Regex.IsMatch(textBox.Text, "[ ^ 0-9]")))
                {
                    textBox.Text = "";
                    textBox.Focus();
                    MessageBox.Show("Неверный номер ТСД!", "Ошибка");
                }
                else prompt.Close();
            };
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(textBox);
            textBox.MaxLength = 2;
            prompt.ShowDialog();
            textBox.Focus();
            return textBox.Text;
        }


        private void registryKeySet()
        {
            try
            {
                RegistryKey rk;
                string key = @"\Software\Symbol\ScanWedge";
                rk = Registry.CurrentUser.CreateSubKey(key);
                rk = Registry.CurrentUser.OpenSubKey(key, true);
                rk.SetValue("AutoEnter", 1);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произвошла ошибка при изменении ключа реестра!", "Ошибка");
            }
        }

        private void scanwedgeStart()
        {
            try
            {
                string path = "";
                if (Directory.Exists(@"\Application\Samples.c"))
                {
                    path = @"\Application\Samples.c\scanwedge.exe";
                }
                else
                {
                    path = @"\Application\scanwedge.exe";
                }
                System.Diagnostics.Process scanWedge = new System.Diagnostics.Process();
                scanWedge.StartInfo.FileName = path;
                bool result = ProcessCE.IsRunning(path);
                if (!result)
                {
                    scanWedge.Start();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при запуске Scanwedge!", "Ошибка");
            }
        }

        private void rdpFileCreate()
        {
            try
            {
                //диалог для запроса от пользователя номера ТСД
                string promptValue = ShowDialog("Номер ТСД", "Введите номер ТСД и нажмите ОК: ");
                
                if (!File.Exists(@"\windows\desktop\1c.rdp")) //проверка папки рабочий стол на наличие rdp-файла
                {
                    //создание файла в случае его отсутствия
                    File.Create(@"\windows\desktop\1c.rdp").Close();
                }

                //открытие файла для записи в него необходимой информации
                StreamWriter swFile =
                             new StreamWriter(
                                new FileStream(@"\windows\desktop\1c.rdp",
                                               FileMode.Truncate),
                                Encoding.Unicode); //кодировку Unicode не менять!

                String output = "";
                if (Directory.Exists(@"\Application\Samples.c"))
                {
                    try
                    {
                        RegistryKey rk;
                        string key = @"\Software\Microsoft\Terminal Server Client\UsernameHint";
                        rk = Registry.CurrentUser.CreateSubKey(key);

                        rk = Registry.CurrentUser.OpenSubKey(key, true);
                        rk.SetValue("server_ip_address_here", @"domain_name_here" + promptValue); //подставить нужные значения адреса и домена
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Произошла ошибка при создании ключа реестра!", "Ошибка");
                    }

                    output = "screen mode id:i:2" + "\r\n" +
                          "span monitors:i:0" + "\r\n" +
                          "desktopwidth:i:320" + "\r\n" +
                          "desktopheight:i:320" + "\r\n" +
                          "session bpp:i:16" + "\r\n" +
                          "winposstr:s:0,1,0,0,320,320" + "\r\n" +
                          "full address:s:*" + "\r\n" + //подставить IP-адрес сервера вместо звёздочки
                          "compression:i:1" + "\r\n" +
                          "keyboardhook:i:2" + "\r\n" +
                          "audiomode:i:0" + "\r\n" +
                          "redirectprinters:i:1" + "\r\n" +
                          "redirectcomports:i:0" + "\r\n" +
                          "redirectsmartcards:i:0" + "\r\n" +
                          "redirectclipboard:i:1" + "\r\n" +
                          "redirectposdevices:i:0" + "\r\n" +
                          "redirectdrives:i:0" + "\r\n" +
                          "displayconnectionbar:i:1" + "\r\n" +
                          "autoreconnection enabled:i:1" + "\r\n" +
                          "authentication level:i:0" + "\r\n" +
                          "prompt for credentials:i:0" + "\r\n" +
                          "negotiate security layer:i:1" + "\r\n" +
                          @"alternate shell:s:""e:\w.exe"" e:\L.vbs" + "\r\n" +
                          "shell working directory:s:" + "\r\n" +
                          "disable wallpaper:i:1" + "\r\n" +
                          "disable full window drag:i:1" + "\r\n" +
                          "allow desktop composition:i:0" + "\r\n" +
                          "allow font smoothing:i:0" + "\r\n" +
                          "disable menu anims:i:1" + "\r\n" +
                          "disable themes:i:0" + "\r\n" +
                          "disable cursor setting:i:0" + "\r\n" +
                          "bitmapcachepersistenable:i:0" + "\r\n" + "";
                }
                else
                {
                    output = "Keyboard Layout:s:00000LOC_DEFAULTLCID" + "\r\n" +
                               "BitmapPersistCacheSize:i:1" + "\r\n" +
                               "BitmapCacheSize:i:21" + "\r\n" +
                               "Shadow Bitmap Enabled:i:1" + "\r\n" +
                               "ColorDepthID:i:3" + "\r\n" +
                               "DesktopHeight:i:324" + "\r\n" +
                               "DesktopWidth:i:324" + "\r\n" +
                               "Disable Themes:i:0" + "\r\n" +
                               "Disable Menu Anims:i:1" + "\r\n" +
                               "Disable Full Window Drag:i:1" + "\r\n" +
                               "Disable Wallpaper:i:1" + "\r\n" +
                               "MaxReconnectAttempts:i:20" + "\r\n" +
                               "KeyboardHookMode:i:1" + "\r\n" +
                               "StartFullScreen:i:1" + "\r\n" +
                               "Compress:i:1" + "\r\n" +
                               "BBarShowPinBtn:i:0" + "\r\n" +
                               "BitmapPersistenceEnabled:i:0" + "\r\n" +
                               "AudioRedirectionMode:i:0" + "\r\n" +
                               "EnablePortRedirection:i:0" + "\r\n" +
                               "EnableDriveRedirection:i:0" + "\r\n" +
                               "AutoReconnectEnabled:i:1" + "\r\n" +
                               "EnableSCardRedirection:i:0" + "\r\n" +
                               "EnablePrinterRedirection:i:1" + "\r\n" +
                               "BBarEnabled:i:1" + "\r\n" +
                               "ServerName:s:*" + "\r\n" + //подставить IP-адрес или имя сервера вместо звёздочки
                               "DisableFileAccess:i:0" + "\r\n" +
                               "MCSPort:i:*" + "\r\n" + //подставить номер порта rdp вместо звёздочки
                               "UserName:s:*" + promptValue + "\r\n" + //подставить имя пользователя (или шаблон имени) на удалённом сервере вместо звёздочки
                               "Domain:s:*" + "\r\n" + //подставить имя домена вместо звёздочки
                               @"AlternateShell:s:""e:\w.exe"" e:\L.vbs" + "\r\n" +
                               "WorkingDir:s:" + "\r\n" + "";
                }

                swFile.Write(output);
                swFile.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при создании удалённого подключения!", "Ошибка");
            }
        }

        private void checkWifiAdapter()
        {
            try
            {
                myConfig = new Config(FusionAccessType.COMMAND_MODE);
                myWlan = new WLAN(FusionAccessType.COMMAND_MODE);
            }
            catch (OperationFailureException ex)
            {
                MessageBox.Show("Ошибка при настройке WiFi", "Ошибка");
            }

            try
            {
                if (myWlan != null)
                {
                    myWlan.Adapters[0].PowerState = Adapter.PowerStates.ON;
                    myWlan.Adapters[0].CountryCode = "RU";
                    myWlan.Adapters[0].IEEE80211dEnabled = false;
                    
                }

            }
            catch (OperationFailureException ex)
            {
                MessageBox.Show("Ошибка при настройке WiFi", "Ошибка");
            }

            myConfig.Dispose();
            myWlan.Dispose();

        }

        private void showStatus()
        {
            try
            {
                string output = "";
                if (File.Exists(@"\windows\desktop\1c.rdp"))
                {
                    output = "Удалённое подключение к 1С создано!" + Environment.NewLine + Environment.NewLine +
                             "Нажмите ОК для выхода из программы.";
                }
                else
                {
                    output = "Удалённое подключение к 1С не создано!" + Environment.NewLine + Environment.NewLine +
                             "Нажмите ОК для выхода из программы.";
                }
                SystemSounds.Hand.Play();
                MessageBox.Show(output, "Статус");
                Application.Exit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка получения статуса", "Ошибка");
                Application.Exit();
            }
        }

    }
}