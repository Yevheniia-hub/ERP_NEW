﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using ERP_NEW.BLL.Interfaces;
using ERP_NEW.BLL.Services;
using ERP_NEW.BLL.DTO;
using ERP_NEW.BLL.DTO.ModelsDTO;
using ERP_NEW.BLL.DTO.SelectedDTO;
using DevExpress.XtraEditors.Controls;
using Ninject;
using System.Web;
using ERP_NEW.BLL.Infrastructure;
using ERP_NEW.BLL;

namespace ERP_NEW.GUI.Accounting
{
    public partial class FixedAssetsOrderAddJournalFm : DevExpress.XtraEditors.XtraForm
    {
        private BindingSource fixedAssetsOrderJournalBS = new BindingSource();
        List<FixedAssetsMaterialsDTO> fixedAssetsOrderMaterialsList=new List<FixedAssetsMaterialsDTO>();
        private BindingSource fixedAssetsOrderArchiveBS = new BindingSource();
        private IFixedAssetsOrderService fixedAssetsOrderService;
                
        private int rezRadioBtn = 0;
        string findTypeOrder = "";
        string rezTab = "";
        private ObjectBase Item
        {
            get { return fixedAssetsOrderJournalBS.Current as ObjectBase; }
            set
            {
                fixedAssetsOrderJournalBS.DataSource = value;
                value.BeginEdit();
            }
        }

        private ObjectBase ItemArchive
        {
            get { return fixedAssetsOrderArchiveBS.Current as ObjectBase; }
            set
            {
                fixedAssetsOrderArchiveBS.DataSource = value;
                value.BeginEdit();
            }
        }
       

        public FixedAssetsOrderAddJournalFm(FixedAssetsOrderJournalDTO model,FixedAssetsOrderArchiveJournalDTO modelArchive ,string rezTagTabPage,List<FixedAssetsMaterialsDTO>modelMaterialsList)
        {
            InitializeComponent();
            fixedAssetsOrderJournalBS.DataSource = model;
            fixedAssetsOrderArchiveBS.DataSource = modelArchive;
            fixedAssetsOrderMaterialsList=modelMaterialsList;
            fixedAssetsOrderService = Program.kernel.Get<IFixedAssetsOrderService>();
            nameFixedAssetsOrder.Text = model.InventoryName;
            dateLabel.Text = model.BeginDate.ToShortDateString();
            rezTab = rezTagTabPage;
        }

        int posNumber;// = (int)
            
        private bool Save()
        {
           
            
            if (numberOrderEdit.Text.Length != 0)
            {
                if (findTypeOrder != "")
                {
                    

                    if (rezRadioBtn != 0)
                    {
                        if (rezTab == "0")//FA
                        {
                            if (rezRadioBtn == 2)//increase fixed assets
                            {
                                //foreach (var item in fixedAssetsOrderMaterialsList)
                                //{
                                //    if (item.Flag == 1)
                                //    {
                                //        yellow = 1;


                                //    }//break; }
                                //    else { yellow = 0; }
                                //}
                                var f= fixedAssetsOrderMaterialsList.FirstOrDefault(a => a.Flag == 1);
                                if (f==null)
                                {
                                    MessageBox.Show("Помилка! Необхідно обрати матеріал, який збільшив вартість основного засобу.", "Увага", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    Close();
                                    return false;
                                }
                                else
                                {

                                    FixedAssetsOrderRegistrationDTO newModel = new FixedAssetsOrderRegistrationDTO()
                                        {
                                            FixedAssetsOrderId = ((FixedAssetsOrderJournalDTO)Item).Id,
                                            NumberOrder = numberOrderEdit.Text,
                                            DateOrder = ((FixedAssetsOrderJournalDTO)Item).BeginDate,//(DateTime)dateEdit1.EditValue,
                                            TypeOrder = findTypeOrder,
                                            StatusTypeOrder = rezRadioBtn,
                                            BeginPrice = ((FixedAssetsOrderJournalDTO)Item).BeginPrice
                                        };
                                    fixedAssetsOrderService.FixedAssetsOrderRegistrationCreate(newModel);
                                    return true;
                                }
                                
                            }

                            else if (rezRadioBtn == 1)//vvedenia
                            {
                                FixedAssetsOrderRegistrationDTO newModel = new FixedAssetsOrderRegistrationDTO()
                                {
                                    FixedAssetsOrderId = ((FixedAssetsOrderJournalDTO)Item).Id,
                                    NumberOrder = numberOrderEdit.Text,
                                    DateOrder = ((FixedAssetsOrderJournalDTO)Item).BeginDate,//(DateTime)dateEdit1.EditValue,
                                    TypeOrder = findTypeOrder,
                                    StatusTypeOrder = rezRadioBtn,
                                    BeginPrice = ((FixedAssetsOrderJournalDTO)Item).BeginPrice
                                };
                                fixedAssetsOrderService.FixedAssetsOrderRegistrationCreate(newModel);
                                return true;
                            }

                            else if (rezRadioBtn == 3)//sale
                            {
                                FixedAssetsOrderRegistrationDTO newModel = new FixedAssetsOrderRegistrationDTO()
                                    {
                                        FixedAssetsOrderId = ((FixedAssetsOrderJournalDTO)Item).Id,
                                        NumberOrder = numberOrderEdit.Text,
                                        DateOrder = ((FixedAssetsOrderJournalDTO)Item).BeginDate,
                                        TypeOrder = findTypeOrder,
                                        StatusTypeOrder = rezRadioBtn,
                                        BeginPrice = ((FixedAssetsOrderJournalDTO)Item).BeginPrice,
                                        BeginDate = ((FixedAssetsOrderJournalDTO)Item).BeginDate,
                                        Supplier_Id = ((FixedAssetsOrderJournalDTO)Item).SupplierId
                                    };
                                fixedAssetsOrderService.FixedAssetsOrderRegistrationCreate(newModel);
                                return true;
                            }
                        }
                            if (rezTab == "1")//archiv
                            {
                                FixedAssetsOrderRegistrationDTO newModelArchive = new FixedAssetsOrderRegistrationDTO()
                                {
                                    NumberOrder = numberOrderEdit.Text,
                                    DateOrder = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).EndRecordDate,
                                    TypeOrder = findTypeOrder,
                                    StatusTypeOrder = rezRadioBtn,
                                    FixedAssetsOrderId = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).Id,
                                    EndRecordDate = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).EndRecordDate,
                                    SoldPrice = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).SoldPrice,
                                    BeginDate = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).BeginDate,
                                    TransferPrice = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).TransferPrice,
                                    Supplier_Id = ((FixedAssetsOrderArchiveJournalDTO)ItemArchive).SupplierId
                                };
                                fixedAssetsOrderService.FixedAssetsOrderRegistrationCreate(newModelArchive);
                                return true;
                            }
                        else return false;
                        }
                        else return false;
                    }
                    else { MessageBox.Show("Тип наказу не обрано!", "", MessageBoxButtons.OK, MessageBoxIcon.Information); return false; }                   
                }
                else return false;
            }      

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (Save())
            {
                Close();
                MessageBox.Show(findTypeOrder+ "№  "+ numberOrderEdit.Text + " успішно створений!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void cancelBtn_Click(object sender, EventArgs e)
        {
            this.Item.EndEdit();
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void orderRadioGroup_SelectedIndexChanged(object sender, EventArgs e)
        {
            RadioGroup edit = sender as RadioGroup;
            
            switch(edit.SelectedIndex)
            {
                case 0: rezRadioBtn= 1;
                    findTypeOrder = "Наказ на введення";
                    break;
                case 1: rezRadioBtn = 2;
                    findTypeOrder = "Наказ на збільшення вартості";
                    break;
                case 2: rezRadioBtn = 3;
                    findTypeOrder = "Наказ на продаж/списання";
                    break;
                default: findTypeOrder = "";
                    rezRadioBtn = 1;
                    break;                    
            }         
        }

        private void orderRadioGroup_FormatEditValue(object sender, ConvertEditValueEventArgs e)
        {
            RadioGroup edit = sender as RadioGroup;
            switch (rezTab)
            {
                case "0":
                    edit.Properties.Items[0].Enabled = true;
                    edit.Properties.Items[1].Enabled = true;
                    edit.Properties.Items[2].Enabled = false;
                    break;
                case "1":
                    edit.Properties.Items[0].Enabled = false;
                    edit.Properties.Items[1].Enabled = false;
                    edit.Properties.Items[2].Enabled = true;
                    break;
            }
        }
    }
}