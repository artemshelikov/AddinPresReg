using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Inventor;
using System.Runtime.InteropServices;
using InvAddIn;
using System.IO;

namespace AddinPresReg
{
    class ButtonMsg
    {

        Inventor.Application m_InvApp;
        //Дефинция кнопки
        Inventor.ButtonDefinition m_ButtonDef;
        //Ribbon-объекты
        ButtonPathToRibbon m_ButtonParhToRibbon;
        //Иконки
        ButtonIcons m_ButtonIcons;
        //ButtonIcons m_ButtonNew;
        Form1 form;

        public ButtonMsg(Inventor.Application InvApp)
        {
            m_InvApp = InvApp;
            //Привязка к событию деактивации AddIn
            StandardAddInServer.RaiseEventAddInDeactivate += OnAddInDeactivate;
            //Создание места на панели куда будет установлена кнопка
            m_ButtonParhToRibbon = new ButtonPathToRibbon(m_InvApp);
            //Создание иконок для кнопки
            m_ButtonIcons = new ButtonIcons(
            InvAddIn.Properties.Resources._01Standart,
            InvAddIn.Properties.Resources.iconBig2, InvAddIn.Properties.Resources.iconBig3);
            //m_ButtonNew = new ButtonIcons(InvAddIn.Properties.Resources.nip,
            //InvAddIn.Properties.Resources.nip);
            //Создание дифинции кнопки и добавление её на панель
            CreateButtonDef_And_InsertToPanel();
        }

        void CreateButtonDef_And_InsertToPanel()
        {
            //Создания дефинции кнопки
            m_ButtonDef = m_InvApp.CommandManager.ControlDefinitions.AddButtonDefinition(
            "Регулятор давления", "MyFirstMessageCmd", Inventor.CommandTypesEnum.kQueryOnlyCmdType,
            "", "Регулятор давления", "Создать", m_ButtonIcons.IconStandard,
            m_ButtonIcons.IconLarge);

            //Редактирование окна при наведении на кнопку
            m_ButtonDef.ProgressiveToolTip.Title = "Регулятор давления";
            m_ButtonDef.ProgressiveToolTip.Description = "Проектирование регулятора давления.";
            m_ButtonDef.ProgressiveToolTip.ExpandedDescription = "Проектирование регулятора давления " +
                "с возможностью изменения параметров деталей в зависимости от необходимости.";
            m_ButtonDef.ProgressiveToolTip.Image = m_ButtonIcons.IconDisplay;

            //Добавление кнопки на панель
            m_ButtonParhToRibbon.RibbonPanelFirst.CommandControls.AddButton(m_ButtonDef,
            true);
            //Привязка к событию клика на кнопку
            m_ButtonDef.OnExecute += OnExecute;
        }
        public Ribbon RibbonZeroDoc1;
        public RibbonTab RibbonTabFirst1;

        void OnExecute(Inventor.NameValueMap Context)
        {
            RibbonZeroDoc1 = m_InvApp.UserInterfaceManager.Ribbons["Assembly"];
            RibbonTabFirst1 = RibbonZeroDoc1.RibbonTabs["id_TabDesign"];
            //System.Windows.Forms.MessageBox.Show("Моя первая кнопка");
            form = new Form1();
            string path = @"C:\Temp\logi.txt";
            StringBuilder res = new StringBuilder();
            foreach (RibbonPanel obj in RibbonTabFirst1.RibbonPanels)
            {
                res.Append($"{obj.InternalName}\n{obj.ClientId}\n");
            }
            string text = String.Join("\n", res);
            using (StreamWriter writer = new StreamWriter(path, false))
            {
                writer.WriteLine(text);
            }
            form.Show();
        }
        

        void OnAddInDeactivate()
        {
            //Отключение событий
            m_ButtonDef.OnExecute -= OnExecute;
            StandardAddInServer.RaiseEventAddInDeactivate -= OnAddInDeactivate;
            //Удаление объектов пользовательского интерфейса
            m_ButtonDef.Delete();
            m_ButtonParhToRibbon.RibbonTabFirst.Delete();
        }

        //Контейнер для Ribbon-объектов
        private class ButtonPathToRibbon
        {
            public readonly Ribbon RibbonZeroDoc;
            public readonly RibbonTab RibbonTabFirst;
            public readonly RibbonPanel RibbonPanelFirst;
            //Конструктор
            public ButtonPathToRibbon(Inventor.Application InvApp)
            {
                RibbonZeroDoc = InvApp.UserInterfaceManager.Ribbons["Assembly"];
                RibbonTabFirst = RibbonZeroDoc.RibbonTabs["id_TabDesign"];
                this.RibbonPanelFirst = RibbonTabFirst.RibbonPanels.Add("Регуляторы", 
                    "id_PanelMessage2", "", "id_PanelA_DesignFrame", true);
                
            }
        }//Конец класса

        //Контейнер-конвертор для иконок
        private class ButtonIcons
        {
            public readonly stdole.IPictureDisp IconStandard;
            public readonly stdole.IPictureDisp IconLarge;
            public readonly stdole.IPictureDisp IconDisplay;
            //Конструктор
            public ButtonIcons(System.Drawing.Image PictureStandard,
            System.Drawing.Image PictireLarge, System.Drawing.Image PictureDisplay)
            {
                IconStandard = ImageConvertor.ConvertImageToIPictureDisp(PictureStandard);
                IconLarge = ImageConvertor.ConvertImageToIPictureDisp(PictireLarge);
                IconDisplay = ImageConvertor.ConvertImageToIPictureDisp(PictureDisplay);
            }
            private class ImageConvertor : System.Windows.Forms.AxHost
            {

                ImageConvertor() : base("") { }
                public static stdole.IPictureDisp ConvertImageToIPictureDisp(
                System.Drawing.Image Image)
                {
                    if (null == Image) return null;
                    return GetIPictureDispFromPicture(Image) as stdole.IPictureDisp;
                }
            }
        }//конец класса



    }
}
