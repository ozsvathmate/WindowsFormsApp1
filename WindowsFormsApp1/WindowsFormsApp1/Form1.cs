﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp; 
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;
        dvd_magyarEntities context = new dvd_magyarEntities();
        List<kolcsonzesek> Kolcsonzesek;

        public Form1()
        {
            InitializeComponent();
            LoadData();
        }

        private void LoadData()
        {
            Kolcsonzesek = context.kolcsonzesek.ToList();
        }


    }
}
