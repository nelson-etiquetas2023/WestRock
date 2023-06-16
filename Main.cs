using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Data.OleDb;
using System.Data;
using excel = Microsoft.Office.Interop.Excel;
using System.Xml.Serialization;
using System.Drawing;
using Microsoft.Reporting.WinForms;
using System.Net.Configuration;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using System.Data.Common;
using WestRockDataPonchesPRO.Libreria;
using System.Web;
using System.Globalization;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using Microsoft.Reporting.Map.WebForms.BingMaps;
//using Microsoft.Office.Interop.Excel;

namespace WestRockDataPonchesPRO
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }
        //cargo lalibreria de horas extras.
        HorasExtras helib = new HorasExtras();
        readonly char CHARACTER_FILL = ' ';
        readonly String CHAR_SEPARATOR = "";
        string DAYS, MONTHS, YEARS, HOURS, MINUTES;
        string filter_devices;
        List<Registro> registros;
        List<Notification> Notifs = new List<Notification>();
        DateTime FechaHora1, FechaHora2;
        readonly string PathAppFolder = @"C:\Reportes";
        Boolean CheckMarkDevice = true;
        string StringConnectEtiqueta = @"Server=DESKTOP-SOA16M2\SQLEXPRESS;Database=BDBioAdminSQL;Trusted_Connection=True";
        DataTable DtShift = new DataTable();
        DataTable DtShiftDetail = new DataTable();
        DataTable DtEmpleados = new DataTable();
        DataTable DtHorasExtras = new DataTable();
        DataTable DtJornadas = new DataTable();

        int shiftid, daysid;

        private void Main_Load(object sender, EventArgs e)
        {
            LoadParameters();
        }
        private void Conn_SQLSERVER()
        {

            try
            {
                //string StringConnectWest = @"Server=172.23.24.82\SQLEXPRESS,1533;Database=BDBioAdminSQL;UID=etiquetas;PWD=123";
                //string StringConnectLaptop = @"Server=STATIONCODE\SQLEXPRESS;Database=BDBioAdminSQL;UID=sa;PWD=Jossycar5%";
                //string StringConnectServer = @"Server=PCDELL;Database=BDBioAdminSQL;UID=netuser;PWD=password";
                //string StringConnectVM = @"Server=DESKTOP-NUMQD8B\SQLEXPRESS;Database=BDBioAdminSQL;UID=netuser;PWD=password";

                SqlConnection conn = new SqlConnection
                {
                    ConnectionString = StringConnectEtiqueta
                };

                SqlCommand comando = new SqlCommand
                {
                    Connection = conn,
                    CommandType = CommandType.Text
                };
                comando.CommandText =
                    comando.CommandText =
                    "SELECT * INTO #MARCAS FROM (SELECT a.IdUser, b.Name," +
                    "a.RecordTime,cast(RecordTime as time) hora_reg," +
                    "a.MachineNumber,b.ExternalReference,b.IdentificationNumber," +
                    "c.Description as departName," +
                    "(select x.ShiftId from [dbo].[UserShift] x where a.IdUser = x.IdUser and a.RecordTime between x.BeginDate and x.EndDate) as ShiftId " +
                    "FROM Record a " +
                    "LEFT JOIN [User] b ON a.IdUser = b.IdUser " +
                    "LEFT JOIN Device c ON a.MachineNumber = c.MachineNumber " +
                    "WHERE (a.RecordTime >= @p1 AND a.RecordTime <= @p2) AND (a.MachineNumber = @p3 OR " +
                    "a.MachineNumber = @p4 OR a.MachineNumber = @p5 OR a.MachineNumber = @p6 OR " +
                    "a.MachineNumber = @p7 OR a.MachineNumber = @p8 OR a.MachineNumber = @p9 OR " +
                    "a.MachineNumber = @p10) AND (b.ExternalReference = CASE WHEN @p11='' " +
                    "THEN b.ExternalReference ELSE @p11 END)) T " +
                    "SELECT t.*,t1.Description as shiftname," +
                    "case when t1.description LIKE '%NOCTURNO%' then 'N' ELSE 'D'end AS type_Shift " +
                    "FROM #MARCAS t " +
                    "left join shift t1 on t1.ShiftId = t.ShiftId";
                SqlParameter p1 = new SqlParameter("@p1", txt_fecha_desde.Value.Date + txt_hour_desde.Value.TimeOfDay);
                SqlParameter p2 = new SqlParameter("@p2", txt_fecha_hasta.Value.Date + txt_hour_hasta.Value.TimeOfDay);
                SqlParameter p3 = new SqlParameter("@p3", chk_dispo1.Checked ? Convert.ToInt32(txt_par1.Text) : 0);
                SqlParameter p4 = new SqlParameter("@p4", chk_dispo2.Checked ? Convert.ToInt32(txt_par2.Text) : 0);
                SqlParameter p5 = new SqlParameter("@p5", chk_dispo3.Checked ? Convert.ToInt32(txt_par3.Text) : 0);
                SqlParameter p6 = new SqlParameter("@p6", chk_dispo4.Checked ? Convert.ToInt32(txt_par4.Text) : 0);
                SqlParameter p7 = new SqlParameter("@p7", chk_dispo5.Checked ? Convert.ToInt32(txt_par5.Text) : 0);
                SqlParameter p8 = new SqlParameter("@p8", chk_dispo6.Checked ? Convert.ToInt32(txt_par6.Text) : 0);
                SqlParameter p9 = new SqlParameter("@p9", chk_dispo7.Checked ? Convert.ToInt32(txt_par7.Text) : 0);
                SqlParameter p10 = new SqlParameter("@p10", chk_dispo8.Checked ? Convert.ToInt32(txt_par8.Text) : 0);
                SqlParameter p11 = new SqlParameter("@p11", txt_IdUser.Text.Trim());
                comando.Parameters.Add(p1);
                comando.Parameters.Add(p2);
                comando.Parameters.Add(p3);
                comando.Parameters.Add(p4);
                comando.Parameters.Add(p5);
                comando.Parameters.Add(p6);
                comando.Parameters.Add(p7);
                comando.Parameters.Add(p8);
                comando.Parameters.Add(p9);
                comando.Parameters.Add(p10);
                comando.Parameters.Add(p11);
                conn.Open();

                SqlDataAdapter da = new SqlDataAdapter();
                DataTable dt = new DataTable();

                da.SelectCommand = comando;
                da.Fill(dt);

                comando.ExecuteNonQuery();
                //conn.Close();

                registros = new List<Registro>();
                int regNUmbers = 0;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    regNUmbers += 1;
                    Registro ponche0 = new Registro
                    {
                        NumberRecord = regNUmbers,
                        UserID = dt.Rows[i]["IdUser"].ToString().PadLeft(10, CHARACTER_FILL),
                        UserName = dt.Rows[i]["Name"].ToString(),
                        Device = dt.Rows[i]["MachineNumber"].ToString().PadRight(10, CHARACTER_FILL),
                        NameDevice = dt.Rows[i]["departName"].ToString(),
                        DateRegistro = dt.Rows[i]["RecordTime"].ToString(),
                        FechaHora_Marca = Convert.ToDateTime(dt.Rows[i]["RecordTime"]),
                        Reference = dt.Rows[i]["ExternalReference"].ToString().PadLeft(10, CHARACTER_FILL),
                        HoraMark = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).ToString("HH:mm tt"),
                        ShiftId = dt.Rows[i]["ShiftId"].ToString(),
                        Departamento = dt.Rows[i]["departName"].ToString(),
                        shiftname = dt.Rows[i]["shiftname"].ToString(),
                        type_Shift = dt.Rows[i]["type_Shift"].ToString()
                    };
                    DAYS = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).Day.ToString().PadLeft(2, '0');
                    MONTHS = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).Month.ToString().PadLeft(2, '0');
                    YEARS = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).Year.ToString().PadLeft(2, '0');
                    HOURS = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).Hour.ToString().PadLeft(2, '0');
                    MINUTES = Convert.ToDateTime(dt.Rows[i]["RecordTime"]).Minute.ToString().PadLeft(2, '0');
                    ponche0.DateRegistro = DAYS + MONTHS + YEARS;
                    ponche0.HourRegistro = HOURS + MINUTES;
                    registros.Add(ponche0);
                }

                GridData.DataSource = registros;
                //formatear el grid.
                GridData.Columns[0].HeaderText = "Fila";
                GridData.Columns[0].Width = 40;
                GridData.Columns[1].HeaderText = "Codigo Empleado";
                GridData.Columns[1].Width = 58;
                GridData.Columns[2].HeaderText = "Nombre del Empleado";
                GridData.Columns[2].Width = 160;
                GridData.Columns[3].HeaderText = "Fecha-Hora Marcaje";
                GridData.Columns[3].Width = 120;
                GridData.Columns[4].HeaderText = "Numero Device";
                GridData.Columns[4].Width = 52;
                GridData.Columns[5].HeaderText = "Descripcion dispositivo";
                GridData.Columns[5].Width = 100;
                GridData.Columns[6].HeaderText = "string fecha";
                GridData.Columns[6].Width = 60;
                GridData.Columns[7].HeaderText = "string hora";
                GridData.Columns[7].Width = 40;
                GridData.Columns[8].HeaderText = "referencia empleado";
                GridData.Columns[8].Width = 56;
                GridData.Columns[9].HeaderText = "ShiftId";
                GridData.Columns[9].Width = 56;
                GridData.Columns[9].HeaderText = "shiftname";
                GridData.Columns[9].Width = 70;
                GridData.Columns[10].HeaderText = "type shift";
                GridData.Columns[10].Width = 70;
                Label_Counts.Text = regNUmbers.ToString() + " Registros Encontrados.";
                if (chk_Import_Excel.Checked == true)
                {
                    DisplayInExcel(registros);
                }
                if (chk_print_report.Checked == true)
                {
                    ReporteGeneral(registros, "REPORTE FULL DATA");
                }
                CreateFileTxtDataProches(registros);
            }
            catch (SqlException ex)
            {
                MessageBox.Show("error al conectar a la base de datos" + ex);
            }
        }
        private void CreateFileTxtDataProches(List<Registro> lista)
        {
            //CREAR EL ARCHIVO TXT.
            using (StreamWriter sw = File.CreateText(Properties.Resources.PATH_FILE_TXT + Properties.Resources.FILENAME))
            {
                foreach (Registro item in lista)
                {
                    sw.WriteLine(item.Device + CHAR_SEPARATOR + item.Reference + CHAR_SEPARATOR + item.DateRegistro + CHAR_SEPARATOR + item.HourRegistro);
                }
            }
            // MUESTRA EL TXT DE TEXTBOX
            using (StreamReader file = new StreamReader(Properties.Resources.PATH_FILE_TXT + Properties.Resources.FILENAME))
            {
                var temp = file.ReadToEnd();
                TXT_TXTARCHIVO.Text = temp;
            }
        }
        private void ReporteGeneral(List<Registro> data, string tipo_repo)
        {
            //definir los parametros a enviar
            ReportParameter[] rParams = new ReportParameter[5]
            {
                new ReportParameter("fromDate", txt_fecha_desde.Text),
                new ReportParameter("toDate", txt_fecha_hasta.Text),
                new ReportParameter("fromHour", txt_hour_desde.Text),
                new ReportParameter("toHour", txt_hour_hasta.Text),
                new ReportParameter("tipo_repo",tipo_repo)
            };
            //definir la data a enviar al reporte
            ReportDataSource rds = new ReportDataSource("DataPonches", data);
            //crear la instancia del vbisualizador de reportes
            ReporteView repoview = new ReporteView();
            repoview.reportViewer1.Reset();
            repoview.reportViewer1.ProcessingMode = ProcessingMode.Local;
            //repoview.reportViewer1.LocalReport.ReportPath = Application.StartupPath + @"\Reports\ReporteGeneral.rdlc";
            repoview.reportViewer1.LocalReport.DataSources.Clear();
            repoview.reportViewer1.LocalReport.DataSources.Add(rds);
            repoview.reportViewer1.LocalReport.SetParameters(rParams);
            repoview.reportViewer1.LocalReport.Refresh();
            repoview.reportViewer1.RefreshReport();
            repoview.Show();
        }
        private void DisplayInExcel(List<Registro> data)
        {
            var ExcelApp = new excel.Application
            {
                Visible = true
            };
            ExcelApp.Workbooks.Add();
            excel._Worksheet workSheet = (excel.Worksheet)ExcelApp.ActiveSheet;
            workSheet.Cells[1, "a"] = "Items";
            workSheet.Cells[1, "b"] = "USer ID";
            workSheet.Cells[1, "c"] = "Nombre del Usuario";
            workSheet.Cells[1, "d"] = "Fecha Registro";
            workSheet.Cells[1, "e"] = "Hora";
            workSheet.Cells[1, "f"] = "Numero Dispositivo";
            workSheet.Cells[1, "g"] = "Descripcion del Dispositivo";
            workSheet.Cells[1, "h"] = "Referencia";
            workSheet.Cells[1, 7].EntireRow.Font.Bold = true;
            var row = 1;
            foreach (var item in data)
            {
                row++;
                workSheet.Cells[row, "A"] = item.NumberRecord;
                workSheet.Cells[row, "B"] = item.UserID;
                workSheet.Cells[row, "C"] = item.UserName;
                workSheet.Cells[row, "D"] = item.FechaHora_Marca;
                workSheet.Cells[row, "E"] = item.HoraMark;
                workSheet.Cells[row, "F"] = item.Device;
                workSheet.Cells[row, "G"] = item.NameDevice;
                workSheet.Cells[row, "H"] = item.Reference;
            }
            ((excel.Range)workSheet.Columns[1]).AutoFit();
            ((excel.Range)workSheet.Columns[2]).AutoFit();
            ((excel.Range)workSheet.Columns[3]).AutoFit();
            ((excel.Range)workSheet.Columns[4]).AutoFit();
            ((excel.Range)workSheet.Columns[5]).AutoFit();
            ((excel.Range)workSheet.Columns[6]).AutoFit();
            ((excel.Range)workSheet.Columns[7]).AutoFit();
            ExcelApp.Worksheets[1].Name = "Registro de Huellas";

        }
        private void CreateXmlGeneric(List<Notification> Notifications)
        {
            XmlSerializer ser = new XmlSerializer(typeof(List<Notification>));
            using (TextWriter wr = new StreamWriter(PathAppFolder + @"\NotifSalidasAreaXml.xml"))
            {
                ser.Serialize(wr, Notifications);
            }
        }
        private void TreeMarklAlarm()
        {
            var TreeMarkDeviceUser = (from q in registros
                                      group q by new { q.UserID, q.Device, q.FechaHora_Marca.Date } into grp
                                      select new
                                      {
                                          Iduser = grp.Key.UserID,
                                          empleado = grp.FirstOrDefault().UserName,
                                          Fecha = grp.Key.Date.ToShortDateString(),
                                          device = grp.Key.Device,
                                          grp.FirstOrDefault().NameDevice,
                                          Marcas = grp.Count()
                                      }).Where(x => x.Marcas > 3).ToList();

            string[] data = new string[5];
            Notifs.Clear();
            foreach (var item in TreeMarkDeviceUser)
            {
                Notification noti = new Notification
                {
                    Tipo = "Notificaciones en los periodos de descanso.",
                    Iduser = item.Iduser.Trim(),
                    Empleado = item.empleado,
                    Device = item.device,
                    NameDevice = item.NameDevice,
                    //Fecha = item.Fecha,
                    Marcaje1 = DateTime.Today.TimeOfDay.ToString(),
                    Marcaje2 = DateTime.Today.TimeOfDay.ToString(),
                    Diffminutes = 0,
                    Marcas = item.Marcas
                };
                Notifs.Add(noti);
                //iterar el grid para resaltar los registro de 3 marcajes.
                foreach (DataGridViewRow row in GridData.Rows)
                {
                    if (row.Cells[1].Value.ToString() == item.Iduser && row.Cells[4].Value.ToString() == item.device)
                    {
                        row.DefaultCellStyle.BackColor = Color.LightBlue;
                    }
                }
                data[0] = item.Iduser;
                data[1] = item.empleado;
                data[2] = item.device;
                data[3] = item.Fecha.ToString();
                data[4] = item.Marcas.ToString();
                list_notification.Items.Add("El usuario con el codigo : " + data[0].Trim() +
                " empleado de nombre : " + data[1].Trim() + " tiene " + item.Marcas + " marcaje en el dispositivo: "
                + data[2].Trim() + " en la fecha: " + data[3].Trim());
            }
        }
        private void ShootNotify900915()
        {
            TimeSpan range900 = new TimeSpan(8, 45, 0);
            TimeSpan range915 = new TimeSpan(10, 45, 0);


            var desca900915 = from q in registros
                              where q.FechaHora_Marca.TimeOfDay >= range900 && q.FechaHora_Marca.TimeOfDay <= range915
                              group q by new { q.UserID, q.FechaHora_Marca.Date } into grp
                              select grp.OrderBy(x => x.UserID).OrderBy(y => y.FechaHora_Marca).Take(2);

            //Colorear las filas de las Marcas de Descanso
            foreach (var grp in desca900915)
            {
                foreach (var item in grp)
                {
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.UserID.Trim() && row.Cells[3].Value.ToString() == item.FechaHora_Marca.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            // Crear las notificaciones
            var noti900915 = (from q in desca900915
                              select new Notification
                              {
                                  Tipo = "Descanso de 9:00am a 9:15am",
                                  Iduser = q.FirstOrDefault().UserID,
                                  Empleado = q.FirstOrDefault().UserName,
                                  Device = q.FirstOrDefault().Device,
                                  Fecha1 = q.FirstOrDefault().FechaHora_Marca,
                                  Fecha2 = q.LastOrDefault().FechaHora_Marca,
                                  Marcaje1 = q.FirstOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Marcaje2 = q.LastOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Diffminutes = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes,
                                  Tardanza = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes - 15
                              }).Where(x => x.Diffminutes >= 15 + Convert.ToDecimal(TXT_RANGEMINUTES_SALIDAS.Text));
            foreach (var item in noti900915)
            {
                list_notification.Items.Add("SOBREPASO DEL TIEMPO DE DESCANSO DEL DESCANSO DE 9:00 AM. - 9:15 AM. NOMBRE DEL TRABAJADOR : " + item.Iduser.Trim()
                    + " " + item.Empleado + " FECHA : " + item.Fecha1.ToShortDateString() +
                    " SALIDA :" + item.Marcaje1 + " ENTRADA : " + item.Marcaje2 + " " + "DURACION DESCANSO (MIN.) : " + item.Diffminutes + " minutos." +
                    " TARDANZA: " + (item.Tardanza).ToString());
            };
            Notifs.Clear();
            Notifs = noti900915.ToList();
        }
        private void ShootNotify12001245()
        {
            TimeSpan range1200 = new TimeSpan(11, 45, 0);
            TimeSpan range1245 = new TimeSpan(13, 30, 0);


            var desca12001245 = from q in registros
                                where q.FechaHora_Marca.TimeOfDay >= range1200 && q.FechaHora_Marca.TimeOfDay <= range1245
                                group q by new { q.UserID, q.FechaHora_Marca.Date } into grp
                                select grp.OrderBy(x => x.UserID).OrderBy(y => y.FechaHora_Marca).Take(2);

            //Colorear las filas de las Marcas de Descanso
            foreach (var grp in desca12001245)
            {
                foreach (var item in grp)
                {
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.UserID.Trim() && row.Cells[3].Value.ToString() == item.FechaHora_Marca.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            // Crear las notificaciones
            var noti12001245 = (from q in desca12001245
                                select new Notification
                                {
                                    Tipo = "Descanso de 12:00pm a 12:45pm",
                                    Iduser = q.FirstOrDefault().UserID,
                                    Empleado = q.FirstOrDefault().UserName,
                                    Device = q.FirstOrDefault().Device,
                                    Fecha1 = q.FirstOrDefault().FechaHora_Marca,
                                    Fecha2 = q.LastOrDefault().FechaHora_Marca,
                                    Marcaje1 = q.FirstOrDefault().FechaHora_Marca.ToLongTimeString(),
                                    Marcaje2 = q.LastOrDefault().FechaHora_Marca.ToLongTimeString(),
                                    Diffminutes = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes,
                                    Tardanza = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes - 45

                                }).Where(x => x.Diffminutes >= 45 + Convert.ToDecimal(TXT_RANGEMINUTES_SALIDAS.Text));
            foreach (var item in noti12001245)
            {
                list_notification.Items.Add("SOBREPASO DEL TIEMPO DE DESCANSO DEL DESCANSO DE 12:00 AM. - 12:45 AM. NOMBRE DEL TRABAJADOR : " + item.Iduser.Trim()
                    + " " + item.Empleado + " FECHA : " + item.Fecha1.ToShortDateString() +
                    " SALIDA :" + item.Marcaje1 + " ENTRADA : " + item.Marcaje2 + " " + "DURACION DESCANSO (MIN.) : " + item.Diffminutes + " minutos." +
                    " TARDANZA: " + (item.Tardanza).ToString());
            };
            foreach (var item in noti12001245)
            {
                Notifs.Add(item);
            }
        }
        private void ShootNotife700745()
        {
            TimeSpan range1900 = new TimeSpan(18, 45, 0);
            TimeSpan range1930 = new TimeSpan(20, 30, 0);


            var desca19001930 = from q in registros
                                where q.FechaHora_Marca.TimeOfDay >= range1900 && q.FechaHora_Marca.TimeOfDay <= range1930
                                group q by new { q.UserID, q.FechaHora_Marca.Date } into grp
                                select grp.OrderBy(x => x.UserID).OrderBy(y => y.FechaHora_Marca).Take(2);

            //Colorear las filas de las Marcas de Descanso
            foreach (var grp in desca19001930)
            {
                foreach (var item in grp)
                {
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.UserID.Trim() && row.Cells[3].Value.ToString() == item.FechaHora_Marca.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            // Crear las notificaciones
            var noti19001930 = (from q in desca19001930
                                select new Notification
                                {
                                    Tipo = "Descaso de 7:00pm a 7:30pm",
                                    Iduser = q.FirstOrDefault().UserID,
                                    Empleado = q.FirstOrDefault().UserName,
                                    Device = q.FirstOrDefault().Device,
                                    Fecha1 = q.FirstOrDefault().FechaHora_Marca,
                                    Fecha2 = q.LastOrDefault().FechaHora_Marca,
                                    Marcaje1 = q.FirstOrDefault().FechaHora_Marca.ToLongTimeString(),
                                    Marcaje2 = q.LastOrDefault().FechaHora_Marca.ToLongTimeString(),
                                    Diffminutes = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes,
                                    Tardanza = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes - 30
                                }).Where(x => x.Diffminutes >= 30 + Convert.ToDecimal(TXT_RANGEMINUTES_SALIDAS.Text));
            foreach (var item in noti19001930)
            {
                list_notification.Items.Add("SOBREPASO DEL TIEMPO DE DESCANSO DEL DESCANSO DE 7:00 PM. - 7:30 PM. NOMBRE DEL TRABAJADOR : " + item.Iduser.Trim()
                    + " " + item.Empleado + " FECHA : " + item.Fecha1.ToShortDateString() +
                    " SALIDA :" + item.Marcaje1 + " ENTRADA : " + item.Marcaje2 + " " + "DURACION DESCANSO (MIN.) : " + item.Diffminutes + " minutos." +
                    " TARDANZA: " + (item.Tardanza).ToString());
            };
            foreach (var item in noti19001930)
            {
                Notifs.Add(item);
            }
        }
        private void ShootNotify900915pm()
        {
            TimeSpan range900 = new TimeSpan(20, 45, 0);
            TimeSpan range915 = new TimeSpan(22, 30, 0);


            var desca900915 = from q in registros
                              where q.FechaHora_Marca.TimeOfDay >= range900 && q.FechaHora_Marca.TimeOfDay <= range915
                              group q by new { q.UserID, q.FechaHora_Marca.Date } into grp
                              select grp.OrderBy(x => x.UserID).OrderBy(y => y.FechaHora_Marca).Take(2);

            //Colorear las filas de las Marcas de Descanso
            foreach (var grp in desca900915)
            {
                foreach (var item in grp)
                {
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.UserID.Trim() && row.Cells[3].Value.ToString() == item.FechaHora_Marca.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            // Crear las notificaciones
            var noti900915 = (from q in desca900915
                              select new Notification
                              {
                                  Tipo = "Descanso de  9:00pm a 9:15m",
                                  Iduser = q.FirstOrDefault().UserID,
                                  Empleado = q.FirstOrDefault().UserName,
                                  Device = q.FirstOrDefault().Device,
                                  Fecha1 = q.FirstOrDefault().FechaHora_Marca,
                                  Fecha2 = q.LastOrDefault().FechaHora_Marca,
                                  Marcaje1 = q.FirstOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Marcaje2 = q.LastOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Diffminutes = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes,
                                  Tardanza = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes - 15
                              }).Where(x => x.Diffminutes >= 15 + Convert.ToDecimal(TXT_RANGEMINUTES_SALIDAS.Text));
            foreach (var item in noti900915)
            {
                list_notification.Items.Add("SOBREPASO DEL TIEMPO DE DESCANSO DEL DESCANSO DE 9:00 PM. - 9:15 PM. NOMBRE DEL TRABAJADOR : " + item.Iduser.Trim()
                    + " " + item.Empleado + " FECHA : " + item.Fecha1.ToShortDateString() +
                    " SALIDA :" + item.Marcaje1 + " ENTRADA : " + item.Marcaje2 + " " + "DURACION DESCANSO (MIN.) : " + item.Diffminutes + " minutos." +
                    " TARDANZA: " + (item.Tardanza).ToString());
            };
            foreach (var item in noti900915)
            {
                Notifs.Add(item);
            }

        }
        private void ShootNotify200230am()
        {
            TimeSpan range200 = new TimeSpan(1, 45, 0);
            TimeSpan range230 = new TimeSpan(3, 30, 0);


            var desca200230 = from q in registros
                              where q.FechaHora_Marca.TimeOfDay >= range200 && q.FechaHora_Marca.TimeOfDay <= range230
                              group q by new { q.UserID, q.FechaHora_Marca.Date } into grp
                              select grp.OrderBy(x => x.UserID).OrderBy(y => y.FechaHora_Marca).Take(2);

            //Colorear las filas de las Marcas de Descanso
            foreach (var grp in desca200230)
            {
                foreach (var item in grp)
                {
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.UserID.Trim() && row.Cells[3].Value.ToString() == item.FechaHora_Marca.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                        }
                    }
                }
            }
            // Crear las notificaciones
            var noti200230 = (from q in desca200230
                              select new Notification
                              {
                                  Tipo = "Descanso de 2:00am a 2:30am",
                                  Iduser = q.FirstOrDefault().UserID,
                                  Empleado = q.FirstOrDefault().UserName,
                                  Device = q.FirstOrDefault().Device,
                                  Fecha1 = q.FirstOrDefault().FechaHora_Marca,
                                  Fecha2 = q.LastOrDefault().FechaHora_Marca,
                                  Marcaje1 = q.FirstOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Marcaje2 = q.LastOrDefault().FechaHora_Marca.ToLongTimeString(),
                                  Diffminutes = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes,
                                  Tardanza = (q.LastOrDefault().FechaHora_Marca.TimeOfDay - q.FirstOrDefault().FechaHora_Marca.TimeOfDay).Minutes - 30
                              }).Where(x => x.Diffminutes >= 30 + Convert.ToDecimal(TXT_RANGEMINUTES_SALIDAS.Text));
            foreach (var item in noti200230)
            {
                list_notification.Items.Add("SOBREPASO DEL TIEMPO DE DESCANSO DEL DESCANSO DE 2:00 AM. - 2:30 AM. NOMBRE DEL TRABAJADOR : " + item.Iduser.Trim()
                    + " " + item.Empleado + " FECHA : " + item.Fecha1.ToShortDateString() +
                    " SALIDA :" + item.Marcaje1 + " ENTRADA : " + item.Marcaje2 + " " + "DURACION DESCANSO (MIN.) : " + item.Diffminutes + " minutos." +
                    " TARDANZA: " + (item.Tardanza).ToString());
            };
            foreach (var item in noti200230)
            {
                Notifs.Add(item);
            }
        }
        private void ShootNotifySalidasArea()
        {
            TimeSpan range730 = new TimeSpan(7, 35, 00);
            TimeSpan range855 = new TimeSpan(8, 55, 00);
            TimeSpan range935 = new TimeSpan(9, 35, 00);
            TimeSpan range1155 = new TimeSpan(11, 55, 00);
            TimeSpan range1250 = new TimeSpan(12, 55, 00);
            TimeSpan range655 = new TimeSpan(18, 55, 00);
            TimeSpan range755 = new TimeSpan(20, 00, 00);
            TimeSpan range2350 = new TimeSpan(23, 50, 00);
            TimeSpan range245 = new TimeSpan(2, 45, 00);
            TimeSpan range720 = new TimeSpan(7, 20, 00);
            var items_filtered = (from q in registros.Where(x => x.FechaHora_Marca.TimeOfDay > range730 && x.FechaHora_Marca.TimeOfDay < range855 ||
                                  x.FechaHora_Marca.TimeOfDay > range935 && x.FechaHora_Marca.TimeOfDay < range1155 || x.FechaHora_Marca.TimeOfDay > range1250
                                  && x.FechaHora_Marca.TimeOfDay < range655 || x.FechaHora_Marca.TimeOfDay > range755 && x.FechaHora_Marca.TimeOfDay < range2350 ||
                                  x.FechaHora_Marca.TimeOfDay > range245 && x.FechaHora_Marca.TimeOfDay < range720)
                                  select q).ToList();
            var items = items_filtered.GroupBy(x => new { x.UserID, x.FechaHora_Marca.Date })
            .Select(x =>
            {
                var sublist = x.OrderBy(y => y.UserID);
                return sublist.Select((item, index) => new
                {
                    index,
                    iduser = item.UserID,
                    device = item.Device,
                    Namedevice = item.NameDevice,
                    username = item.UserName,
                    salida = index == 0 ? item.FechaHora_Marca : items_filtered.ElementAt(index - 1).FechaHora_Marca,
                    entrada = item.FechaHora_Marca,
                    diff = (item.FechaHora_Marca - (index == 0 ? item.FechaHora_Marca : items_filtered.ElementAt(index - 1).FechaHora_Marca)).Minutes
                }).Where(t => t.index % 2 != 0)
                  .Where(t => t.diff > Convert.ToInt32(TXT_RANGEMINUTES_SALIDAS.Text));
            }).ToList();
            Notifs.Clear();
            foreach (var grp in items)
            {
                foreach (var item in grp)
                {
                    Notification noti = new Notification
                    {
                        Tipo = "Salidas del area de TRabajo.",
                        Iduser = item.iduser.Trim(),
                        Empleado = item.username,
                        Device = item.device,
                        NameDevice = item.Namedevice,
                        //Fecha = item.salida.ToShortDateString(),
                        Marcaje1 = item.salida.ToLongTimeString(),
                        Marcaje2 = item.entrada.ToLongTimeString(),
                        Diffminutes = item.diff
                    };
                    Notifs.Add(noti);
                    // agrego las notificaciones.
                    list_notification.Items.Add(item.salida.Date.ToShortDateString() + " El usuario: " + item.iduser.Trim() + " " + item.username.Trim()
                        + " tuvo una salida del area de trabajo a la hora : " + item.salida.TimeOfDay + " regreso a las : " + item.entrada.TimeOfDay + ". Duracion : " + item.diff
                        + " minutos ");
                    //resaldo los resgistos en el grid.
                    foreach (DataGridViewRow row in GridData.Rows)
                    {
                        if (row.Cells[1].Value.ToString().Trim() == item.iduser.Trim() && row.Cells[3].Value.ToString() == item.salida.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightSkyBlue;
                        }
                        if (row.Cells[1].Value.ToString().Trim() == item.iduser.Trim() && row.Cells[3].Value.ToString() == item.entrada.ToString())
                        {
                            row.DefaultCellStyle.BackColor = Color.LightSkyBlue;
                        }
                    }
                }
            }

        }
        private void BorrarListNotification()
        {
            list_notification.Items.Clear();
            foreach (DataGridViewRow row in GridData.Rows)
            {
                row.DefaultCellStyle.BackColor = Color.White;
            }
        }
        private void Btn_clear_notificaction_Click(object sender, EventArgs e)
        {
            BorrarListNotification();
        }
        private void ReporteNotificaciones(string source, string reportname)
        {
            //definir los parametros a enviar
            ReportParameter[] rParams = new ReportParameter[4]
            {
                new ReportParameter("fromDate", txt_fecha_desde.Text),
                new ReportParameter("toDate", txt_fecha_desde.Text),
                new ReportParameter("fromHour", txt_hour_desde.Text),
                new ReportParameter("toHour", txt_hour_hasta.Text),
            };
            //definir la data a enviar al reporte
            ReportDataSource rds = new ReportDataSource(source, Notifs);
            //crear la instancia del vbisualizador de reportes
            ReporteView repoview = new ReporteView();
            repoview.reportViewer1.Reset();
            repoview.reportViewer1.ProcessingMode = ProcessingMode.Local;
            //repoview.reportViewer1.LocalReport.ReportPath = Application.StartupPath + reportname;
            repoview.reportViewer1.LocalReport.DataSources.Clear();
            repoview.reportViewer1.LocalReport.DataSources.Add(rds);
            repoview.reportViewer1.LocalReport.SetParameters(rParams);
            repoview.reportViewer1.LocalReport.Refresh();
            repoview.reportViewer1.RefreshReport();
            repoview.Show();
        }
        private void Btn_imprimir_Click(object sender, EventArgs e)
        {
            //if (RA_SALIDAS_AREA.Checked) 
            //{
            //    ReporteNotificaciones("WorkArea", @"\Reports\Salidas_WorkArea.rdlc");
            //}
            if (RA_PERMISOS_NOTIFY.Checked)
            {
                ReporteNotificaciones("Notificaciones", @"\Reports\Report_NotiDescansos.rdlc");
            }
            if (RA_MARK_MULTIPLE.Checked)
            {
                ReporteNotificaciones("Data_MultipleMark", @"\Reports\ReportMultipleMark.rdlc");
            }
        }

        private void Bot_handlerEvents_Click(object sender, EventArgs e)
        {
            if (GridData.Rows.Count == 0)
            {
                MessageBox.Show("No hay data Cargada...");
                return;
            }
            BorrarListNotification();
            //if (RA_SALIDAS_AREA.Checked) 
            //{
            //    ShootNotifySalidasArea();
            //}
            if (RA_PERMISOS_NOTIFY.Checked)
            {
                ShootNotify900915();
                ShootNotify12001245();
                ShootNotife700745();
                ShootNotify900915pm();
                ShootNotify200230am();
            }
            if (RA_MARK_MULTIPLE.Checked)
            {
                TreeMarklAlarm();
            }
        }
        private void Link_UnMarkAll_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            chk_dispo1.Checked = false;
            chk_dispo2.Checked = false;
            chk_dispo3.Checked = false;
            chk_dispo4.Checked = false;
            chk_dispo5.Checked = false;
            chk_dispo6.Checked = false;
            chk_dispo7.Checked = false;
            chk_dispo8.Checked = false;
        }

        private void Link_UnMarkAll_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (CheckMarkDevice == true)
            {
                chk_dispo1.Checked = false;
                chk_dispo2.Checked = false;
                chk_dispo3.Checked = false;
                chk_dispo4.Checked = false;
                chk_dispo5.Checked = false;
                chk_dispo6.Checked = false;
                chk_dispo7.Checked = false;
                chk_dispo8.Checked = false;
                CheckMarkDevice = false;
            }
            else
            {
                chk_dispo1.Checked = true;
                chk_dispo2.Checked = true;
                chk_dispo3.Checked = true;
                chk_dispo4.Checked = true;
                chk_dispo5.Checked = true;
                chk_dispo6.Checked = true;
                chk_dispo7.Checked = true;
                chk_dispo8.Checked = true;
                CheckMarkDevice = true;
            }

        }

        private void BOT_LOAD_HORARIOS_Click(object sender, EventArgs e)
        {
            LoadDataHorarios();
        }
        private void LoadDataHorarios()
        {
            //comando para llamar el nombre de los horarios
            SqlConnection conn = new SqlConnection(StringConnectEtiqueta);
            SqlCommand comando = new SqlCommand
            {
                Connection = conn,
                CommandType = CommandType.Text
            };
            comando.CommandText = comando.CommandText =
            "SELECT a.ShiftId,a.Description,a.Comment,a.Cycle FROM Shift a";
            conn.Open();
            comando.ExecuteNonQuery();

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = comando;
            da.Fill(DtShift);



            CBO_HORARIOS.DataSource = DtShift;
            CBO_HORARIOS.DisplayMember = "Description";
            CBO_HORARIOS.ValueMember = "ShiftId";

        }

        private void LoadDetailsShift(int shiftid, int daysid)
        {
            DtShiftDetail.Clear();
            //comando para llamar el nombre de los detalles-horarios
            SqlConnection conn = new SqlConnection(StringConnectEtiqueta);
            SqlCommand comando = new SqlCommand
            {
                Connection = conn,
                CommandType = CommandType.Text
            };
            comando.CommandText = comando.CommandText =
            "SELECT b.IdUser,c.Name,a.ShiftId,DayId,description,type,t2inhour,t2outhour" +
            ",t2overtime1beginhour,t2overtime1endhour,t2overtime1factor," +
            "t2overtime2beginhour,t2overtime2endhour,t2overtime2factor," +
            "t2overtime3beginhour,t2overtime3endhour,t2overtime3factor," +
            "t2overtime4beginhour,t2overtime4endhour,t2overtime4factor," +
            "t2overtime5beginhour,t2overtime5endhour,t2overtime5factor " +
            "FROM[BDBioAdminSQL].[dbo].[ShiftDetail] a " +
            "left join[BDBioAdminSQL].[dbo].[UserShift] b on a.ShiftId = b.ShiftId " +
            "left join[BDBioAdminSQL].[dbo].[User] c on c.IdUser = b.IdUser " +
            "where a.ShiftId = @p1 and DayId = @p2 order by DayId";
            SqlParameter p1 = new SqlParameter("@p1", shiftid);
            SqlParameter p2 = new SqlParameter("@p2", daysid);
            comando.Parameters.Add(p1);
            comando.Parameters.Add(p2);
            conn.Open();
            comando.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = comando;

            da.Fill(DtShiftDetail);
            TXT_HORARIO_START.Text = DtShiftDetail.Rows[0]["t2inhour"].ToString();
            TXT_HORARIOS_FINISH.Text = DtShiftDetail.Rows[0]["t2outhour"].ToString();
            if (DtShiftDetail.Rows[0]["type"].ToString() == "1")
            {
                TXT_HORARIOS_TYPES.Text = "UN MINIMO DE HORAS SIN IMPORTAR LAS ENTRADAS NI LAS SALIDAS.";
            }
            else
            {
                TXT_HORARIOS_TYPES.Text = "HORARIO FIJO CON ENTRADAS Y SALIDAS.";
            }
            if (DtShiftDetail.Rows.Count == 0)
            {
                MessageBox.Show("no hay data");
            }

            // horas extras calc 1
            TXT_HE_LAPSO1A.Text = DtShiftDetail.Rows[0]["t2overtime1beginhour"].ToString();
            TXT_HE_LAPSO1B.Text = DtShiftDetail.Rows[0]["t2overtime1endhour"].ToString();
            TXT_PROCENTAJE1.Text = DtShiftDetail.Rows[0]["t2overtime1factor"].ToString();
            // horas extras calc 2
            TXT_HE_LAPSO2A.Text = DtShiftDetail.Rows[0]["t2overtime2beginhour"].ToString();
            TXT_HE_LAPSO2B.Text = DtShiftDetail.Rows[0]["t2overtime2endhour"].ToString();
            TXT_PROCENTAJE2.Text = DtShiftDetail.Rows[0]["t2overtime2factor"].ToString();
            // horas extras calc 3
            TXT_HE_LAPSO3A.Text = DtShiftDetail.Rows[0]["t2overtime3beginhour"].ToString();
            TXT_HE_LAPSO3B.Text = DtShiftDetail.Rows[0]["t2overtime3endhour"].ToString();
            TXT_PROCENTAJE3.Text = DtShiftDetail.Rows[0]["t2overtime3factor"].ToString();
            // horas extras calc 4
            TXT_HE_LAPSO4A.Text = DtShiftDetail.Rows[0]["t2overtime4beginhour"].ToString();
            TXT_HE_LAPSO4B.Text = DtShiftDetail.Rows[0]["t2overtime4endhour"].ToString();
            TXT_PROCENTAJE4.Text = DtShiftDetail.Rows[0]["t2overtime4factor"].ToString();
            // horas extras calc 5
            TXT_HE_LAPSO5A.Text = DtShiftDetail.Rows[0]["t2overtime5beginhour"].ToString();
            TXT_HE_LAPSO5B.Text = DtShiftDetail.Rows[0]["t2overtime5endhour"].ToString();
            TXT_PROCENTAJE5.Text = DtShiftDetail.Rows[0]["t2overtime5factor"].ToString();

        }


        private void CBO_HORARIOS_SelectedIndexChanged(object sender, EventArgs e)
        {
            TXT_HORARIO_ID.Text = CBO_HORARIOS.SelectedValue.ToString();
        }

        private void bot_cargar_empleado_Click(object sender, EventArgs e)
        {
            grid_empleados.DataSource = helib.GetEmpleados();
        }
        private void CALCULO_HORAS_EXTRAS()
        {
            if (registros == null)
            {
                MessageBox.Show("Cargue los datos de los ponches...");
                return;
            }
            //definir las columnas del reporte
            if (DtHorasExtras.Columns.Count == 0)
            {
                DtHorasExtras.Columns.Add("UserId", typeof(int));
                DtHorasExtras.Columns.Add("UserName", typeof(string));
                DtHorasExtras.Columns.Add("Jornada", typeof(string));
                DtHorasExtras.Columns.Add("Departamento", typeof(string));
                DtHorasExtras.Columns.Add("FechaMarca", typeof(DateTime));
                DtHorasExtras.Columns.Add("HorasExtras", typeof(decimal));
                DtHorasExtras.Columns.Add("Factor", typeof(int));
                DtHorasExtras.Columns.Add("Salario", typeof(double));
                DtHorasExtras.Columns.Add("SalarioFraccion", typeof(double));
                DtHorasExtras.Columns.Add("Monto", typeof(double));
            }
            DtHorasExtras.Clear();
            //Recorrer todos los ponches.
            foreach (var item in registros)
            {
                //buscar los parametros del horario fijo.
                int userid = Convert.ToInt16(item.UserID);
                int shiftf = helib.ObtenerHorarioEmpleado(userid, Convert.ToDateTime(item.DateRegistro));
                int indexday = Convert.ToInt16(item.FechaHora_Marca.DayOfWeek - 1);
                if (indexday == -1) indexday = 6;
                string dia = item.FechaHora_Marca.ToString("dddd");
                HORARIO_FIJO hfp = helib.ObtenerParametrosHorarios(shiftf, indexday);
                item.Jornada = "de: " + hfp.Jornada_Start + " a: " + hfp.Jornada_End;
                item.Departamento = helib.GetDepartamento(userid);
                //nivel 1 de condiciones
                var ch1_str1 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal1_Start + " :00:00 PM";
                var ch1_str2 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal1_End + " :00:00 PM";
                DateTime ch1_date1 = Convert.ToDateTime(ch1_str1);
                DateTime ch1_date2 = Convert.ToDateTime(ch1_str2);
                int ch1_factor = Convert.ToInt16(hfp.Cal1_Factor);
                TimeSpan ch1 = ch1_date2 - ch1_date1;
                //nivel 2 de condiciones
                var ch2_str1 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal2_Start + " :00:00 PM";
                var ch2_str2 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal2_End + " :00:00 PM";
                DateTime ch2_date1 = Convert.ToDateTime(ch2_str1);
                DateTime ch2_date2 = Convert.ToDateTime(ch2_str2);
                int ch2_factor = Convert.ToInt16(hfp.Cal2_Factor);
                TimeSpan ch2 = ch2_date2 - ch2_date1;
                //nivel 3 de condiciones
                var ch3_str1 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal3_Start + " :00:00 PM";
                var ch3_str2 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal3_End + " :00:00 PM";
                DateTime ch3_date1 = Convert.ToDateTime(ch3_str1);
                DateTime ch3_date2 = Convert.ToDateTime(ch3_str2);
                int ch3_factor = Convert.ToInt16(hfp.Cal3_Factor);
                TimeSpan ch3 = ch3_date2 - ch3_date1;
                //nivel 4 de condiciones
                var ch4_str1 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal4_Start + " :00:00 PM";
                var ch4_str2 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal4_End + " :00:00 PM";
                DateTime ch4_date1 = Convert.ToDateTime(ch4_str1);
                DateTime ch4_date2 = Convert.ToDateTime(ch4_str2);
                int ch4_factor = Convert.ToInt16(hfp.Cal4_Factor);
                TimeSpan ch4 = ch4_date2 - ch4_date1;
                //nivel 5 de condiciones
                var ch5_str1 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal5_Start + " :00:00 PM";
                var ch5_str2 = item.FechaHora_Marca.ToShortDateString() + " " + hfp.Cal5_End + " :59:59 PM";
                DateTime ch5_date1 = Convert.ToDateTime(ch5_str1);
                DateTime ch5_date2 = Convert.ToDateTime(ch5_str2);
                int ch5_factor = Convert.ToInt16(hfp.Cal5_Factor);
                TimeSpan ch5 = ch5_date2 - ch5_date1;
                //hay horas extras que calcular
                if (item.FechaHora_Marca > ch1_date1)
                {
                    TimeSpan het = item.FechaHora_Marca - ch1_date1;
                    Boolean run = false;
                    //Obtener el salario x Hora.
                    double sh = helib.ObtenerSalarioxHora(userid);
                    // Condiciones del 1er. nivel.
                    if (het >= ch1)
                    {
                        DataRow row = DtHorasExtras.NewRow();
                        row["UserId"] = item.UserID;
                        row["UserName"] = item.UserName;
                        row["Departamento"] = item.Departamento;
                        row["Jornada"] = item.Jornada;
                        row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                        double horas_extras = (ch1.TotalMinutes / 60);
                        row["HorasExtras"] = horas_extras;
                        row["Factor"] = ch1_factor;
                        row["Salario"] = sh;
                        double salario_factor = (sh * ch1_factor) / 100;
                        row["SalarioFraccion"] = salario_factor;
                        row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                        DtHorasExtras.Rows.Add(row);
                        run = true;
                    }
                    else
                    {
                        DataRow row = DtHorasExtras.NewRow();
                        row["UserId"] = item.UserID;
                        row["UserName"] = item.UserName;
                        row["Departamento"] = item.Departamento;
                        row["Jornada"] = item.Jornada;
                        row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                        double horas_extras = Math.Round((het.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                        row["HorasExtras"] = horas_extras;
                        row["Factor"] = ch1_factor;
                        row["Salario"] = sh;
                        double salario_factor = (sh * ch1_factor) / 100;
                        row["SalarioFraccion"] = salario_factor;
                        row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                        DtHorasExtras.Rows.Add(row);
                        run = false;
                    }
                    // Condiciones de 2do. Nivel.
                    if (run)
                    {
                        TimeSpan het1 = item.FechaHora_Marca - ch1_date2;
                        if (het1 >= ch2)
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((ch2.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch2_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = true;
                        }
                        else
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((het1.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch2_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = false;
                        }
                    }

                    // Condiciones de 3er. Nivel.
                    if (run)
                    {
                        TimeSpan het2 = item.FechaHora_Marca - ch3_date1;
                        if (het2 >= ch3)
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((ch3.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["Jornada"] = item.Jornada;
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch3_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = true;
                        }
                        else
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((het2.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch3_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = false;
                        }
                    }

                    // Condiciones de 4to. Nivel.
                    if (run)
                    {
                        TimeSpan het3 = item.FechaHora_Marca - ch4_date1;
                        if (het3 >= ch4)
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((ch4.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["Jornada"] = item.Jornada;
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch4_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch4_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = true;
                        }
                        else
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((het3.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch4_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch4_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = false;
                        }
                    }

                    // Condiciones de 5to. Nivel.
                    if (run)
                    {
                        TimeSpan het4 = item.FechaHora_Marca - ch5_date1;
                        if (het4 >= ch5)
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((ch5.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["Jornada"] = item.Jornada;
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch5_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch5_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = true;
                        }
                        else
                        {
                            DataRow row = DtHorasExtras.NewRow();
                            row["UserId"] = item.UserID;
                            row["UserName"] = item.UserName;
                            row["Departamento"] = item.Departamento;
                            row["Jornada"] = item.Jornada;
                            row["FechaMarca"] = item.FechaHora_Marca.ToShortDateString();
                            double horas_extras = Math.Round((het4.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras"] = horas_extras;
                            row["Factor"] = ch5_factor;
                            row["Salario"] = sh;
                            double salario_factor = (sh * ch5_factor) / 100;
                            row["SalarioFraccion"] = salario_factor;
                            row["Monto"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            DtHorasExtras.Rows.Add(row);
                            run = false;
                        }
                    }
                }
                Grid_HorasExtras.DataSource = DtHorasExtras;
            }
        }

        private void BOT_CALCULAR_HORASEXTRAS_Click(object sender, EventArgs e)
        {

            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            helib.Repo_HorasExtras_Detalle(txt_fecha_desde.Text, txt_fecha_hasta.Text, txt_hour_desde.Text,txt_hour_hasta.Text, DtHorasExtras);
        }

        private void btn_ejecutar_Click_1(object sender, EventArgs e)
        {
            if (chk_process_data.Checked && GridData.Rows.Count == 0)
            {
                MessageBox.Show("Para filtrar la informacion debe cargar la data primero...");
                return;
            }
            if (RA_ACCESS.Checked)
            {
                Conn_AccessDB();
            }
            if (RA_SQLSERVER.Checked && chk_process_data.Checked == false)
            {
                Conn_SQLSERVER();
            }
            if (chk_process_data.Checked)
            {
                ProcessDataFilter();
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Grid_HorasExtras_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void bot_horas_Click(object sender, EventArgs e)
        {
           


        }

        private void Calculo_Horas_Jornadas() 
        {
            if (registros == null)
            {
                MessageBox.Show("Cargue los datos de los ponches...");
                return;
            }
            //Consulta linq para convertir los registros de los ponches en columnas.


            TimeSpan Entrada_Horario = new TimeSpan(18,0,0);
            TimeSpan Salida_Horario = new TimeSpan(6, 0, 0);
                      
            TimeSpan tsresult = Entrada_Horario.Subtract(Salida_Horario);


            var jornadas = (from q in registros
                            group q by new { q.UserID, fecha = q.type_Shift.Equals("N") ? 
                                             q.FechaHora_Marca.AddHours(15).Date : 
                                             q.FechaHora_Marca.Date } into grp
                            select new
                            {
                                IdUser = grp.Key.UserID,
                                Empleado = grp.FirstOrDefault().UserName,
                                Fecha = grp.FirstOrDefault().FechaHora_Marca,
                                Mark1 = Convert.ToString(grp.FirstOrDefault().FechaHora_Marca.ToShortTimeString()),
                                Mark2 = grp.Count() > 1 ? Convert.ToString(grp.ElementAtOrDefault(1).FechaHora_Marca.ToShortTimeString()) : "",
                                Mark3 = grp.Count() > 2 ? Convert.ToString(grp.ElementAtOrDefault(2).FechaHora_Marca.ToShortTimeString()) : "",
                                Mark4 = grp.Count() > 3 ? Convert.ToString(grp.ElementAtOrDefault(3).FechaHora_Marca.ToShortTimeString()) : "",
                                Ponches = grp.Count(),
                                Horas_Jornada = (grp.LastOrDefault().FechaHora_Marca - grp.FirstOrDefault().FechaHora_Marca)
                            }).ToList();

            //definir las columnas del reporte.
            if (DtJornadas.Columns.Count == 0)
            {
                DtJornadas.Columns.Add("UserId", typeof(int));
                DtJornadas.Columns.Add("Empleado", typeof(string));
                DtJornadas.Columns.Add("Fecha", typeof(DateTime));
                DtJornadas.Columns.Add("Horario", typeof(string));
                DtJornadas.Columns.Add("Horario_Entrada", typeof(string));
                DtJornadas.Columns.Add("Horario_Salida", typeof(string));
                DtJornadas.Columns.Add("Mark1", typeof(string));
                DtJornadas.Columns.Add("Mark2", typeof(string));
                DtJornadas.Columns.Add("Mark3", typeof(string));
                DtJornadas.Columns.Add("Mark4", typeof(string));
                DtJornadas.Columns.Add("TardanzaEntrada", typeof(double));
                DtJornadas.Columns.Add("Ponches", typeof(Int32));
                DtJornadas.Columns.Add("Horas_Jornada", typeof(double));
                DtJornadas.Columns.Add("Horas_Extras", typeof(TimeSpan));
                DtJornadas.Columns.Add("HorasExtras1", typeof(double));
                DtJornadas.Columns.Add("Factor1", typeof(Int32));
                DtJornadas.Columns.Add("SalarioHora", typeof(double));
                DtJornadas.Columns.Add("SalarioFraccion1", typeof(double));
                DtJornadas.Columns.Add("MontoExtra1", typeof(double));
                DtJornadas.Columns.Add("HorasExtras2", typeof(double));
                DtJornadas.Columns.Add("Factor2", typeof(Int32));
                DtJornadas.Columns.Add("SalarioFraccion2", typeof(double));
                DtJornadas.Columns.Add("MontoExtra2", typeof(double));
            }
            DtJornadas.Clear();
            //Recorrer todos los ponches.
            foreach (var item in jornadas)
            {
                int userid = Convert.ToInt16(item.IdUser);
                int shiftf = helib.ObtenerHorarioEmpleado(userid,Convert.ToDateTime(item.Fecha));
                int indexday = Convert.ToInt16(Convert.ToDateTime(item.Fecha).DayOfWeek - 1);
                if (indexday == -1) indexday = 6;
                string dia = Convert.ToDateTime(item.Fecha).ToString("dddd");
                HORARIO_FIJO hfp = helib.ObtenerParametrosHorarios(shiftf, indexday);
                //LLenar el datatable con los parametros de los horarios.
                DataRow row = DtJornadas.NewRow();
                row["UserId"] = item.IdUser;
                row["Empleado"] = item.Empleado;
                row["Horario"] = hfp.Horario_Name;
                var Jor_start = item.Fecha.ToShortDateString() + " " + hfp.Jornada_Start + ":00:00";
                var Jor_end = item.Fecha.ToShortDateString()   + " " + hfp.Jornada_End   + ":00:00";
                // calculo horas horario
                DateTime date_jorstart = Convert.ToDateTime(Jor_start);
                DateTime date_jorend  = Convert.ToDateTime(Jor_end);
                int horas_horario =  (date_jorend - date_jorstart).Hours < 0 ? ((date_jorend - date_jorstart).Hours)*-1 : (date_jorend - date_jorstart).Hours;
                row["Horario_Entrada"] = hfp.Jornada_Start;
                row["Horario_Salida"] = hfp.Jornada_End;
                // calculo de tardanza
                var tar_entrada = item.Fecha - Convert.ToDateTime(Jor_start) < TimeSpan.Zero ? TimeSpan.Zero : item.Fecha - Convert.ToDateTime(Jor_start);
                row["TardanzaEntrada"] = tar_entrada.Minutes;
                row["Fecha"] = item.Fecha;
                row["Mark1"] = item.Mark1;
                row["Mark2"] = item.Mark2;
                row["Mark3"] = item.Mark3;
                row["Mark4"] = item.Mark4;
                row["Ponches"] = item.Ponches;
                double horas_jor = Math.Round((item.Horas_Jornada.TotalHours), 2, MidpointRounding.AwayFromZero);
                row["Horas_Jornada"] = horas_jor;
                //Verificacion de condicion de jornada con pronches 2 dias.
                string jor_start_ampm = date_jorstart.ToString("tt", CultureInfo.InvariantCulture);
                string col1_mark1_ampm = item.Fecha.ToString("tt", CultureInfo.InvariantCulture);

               


                // Calculo Datetime Marcaje Salida Empleado.
                var SalidaStr = item.Fecha.ToShortDateString();
                DateTime SalidaJornada = Convert.ToDateTime(SalidaStr +" "+ item.Mark4);
                //nivel 1 de condiciones
                var ch1_str1 = SalidaJornada.ToShortDateString() + " " + hfp.Cal1_Start + " :00:00 PM";
                var ch1_str2 = SalidaJornada.ToShortDateString() + " " + hfp.Cal1_End   + " :00:00 PM";
                DateTime ch1_date1 = Convert.ToDateTime(ch1_str1);
                DateTime ch1_date2 = Convert.ToDateTime(ch1_str2);
                int ch1_factor = Convert.ToInt16(hfp.Cal1_Factor);
                TimeSpan ch1 = ch1_date2 - ch1_date1;
                //nivel 2 de condiciones
                var ch2_str1 = SalidaJornada.ToShortDateString() + " " + hfp.Cal2_Start + " :00:00 PM";
                var ch2_str2 = SalidaJornada.ToShortDateString() + " " + hfp.Cal2_End   + " :00:00 PM";
                DateTime ch2_date1 = Convert.ToDateTime(ch2_str1);
                DateTime ch2_date2 = Convert.ToDateTime(ch2_str2);
                int ch2_factor = Convert.ToInt16(hfp.Cal2_Factor);
                TimeSpan ch2 = ch2_date2 - ch2_date1;
                //Calculo Horas Extras Empleado
                row["Horas_Extras"] = (SalidaJornada - ch1_date1) < TimeSpan.Zero ? TimeSpan.Zero : (SalidaJornada - ch1_date1);
                //hay horas extras que calcular
                if (SalidaJornada > ch1_date1)
                {
                    TimeSpan het = SalidaJornada - ch1_date1;
                    Boolean run = false;
                    //Obtener el salario x Hora.
                    double sh = helib.ObtenerSalarioxHora(userid);
                    row["SalarioHora"] = sh;
                    // Condiciones del 1er. nivel.
                    if (het >= ch1)
                    {
                        double horas_extras = (ch1.TotalMinutes / 60);
                        row["HorasExtras1"] = horas_extras;
                        row["Factor1"] = ch1_factor;
                        double salario_factor = (sh * ch1_factor) / 100;
                        row["SalarioFraccion1"] = salario_factor;
                        row["MontoExtra1"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                        run = true;
                    }
                    else
                    {
                        double horas_extras = Math.Round((het.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                        row["HorasExtras1"] = horas_extras;
                        row["Factor1"] = ch1_factor;
                        double salario_factor = (sh * ch1_factor) / 100;
                        row["SalarioFraccion1"] = salario_factor;
                        row["MontoExtra1"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                        run = false;
                    }
                    // Condiciones de 2do. Nivel.
                    if (run)
                    {
                        TimeSpan het1 = SalidaJornada - ch1_date2;
                        if (het1 >= ch2)
                        {
                            double horas_extras = Math.Round((ch2.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras2"] = horas_extras;
                            row["Factor2"] = ch2_factor;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion2"] = salario_factor;
                            row["MontoExtra2"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            run = true;
                        }
                        else
                        {
                            double horas_extras = Math.Round((het1.TotalMinutes / 60), 2, MidpointRounding.AwayFromZero);
                            row["HorasExtras2"] = horas_extras;
                            row["Factor2"] = ch2_factor;
                            double salario_factor = (sh * ch2_factor) / 100;
                            row["SalarioFraccion2"] = salario_factor;
                            row["MontoExtra2"] = Math.Round((salario_factor * horas_extras), 2, MidpointRounding.AwayFromZero);
                            run = false;
                        }
                    }
                }
                DtJornadas.Rows.Add(row);
            }
            Grid_HorasExtras.DataSource = DtJornadas;
        }

        private void BETN_REPO_JORJNADAS_Click(object sender, EventArgs e)
        {
            Calculo_Horas_Jornadas();
            helib.Repo_Horas_Jornada(txt_fecha_desde.Text, txt_fecha_hasta.Text, txt_hour_desde.Text, txt_hour_hasta.Text, DtJornadas);
        }

        private void BOT_CARGAR_DIA_Click(object sender, EventArgs e)
        {
            shiftid = Convert.ToInt16(TXT_HORARIO_ID.Text);
            if (RAD_LUNES.Checked)
            {
                daysid = 0;
            }
            if (RAD_MARTES.Checked)
            {
                daysid = 1;
            }
            if (RAD_MIERCOLES.Checked)
            {
                daysid = 2;
            }
            if (RAD_JUEVES.Checked)
            {
                daysid = 3;
            }
            if (RAD_VIERNES.Checked)
            {
                daysid = 4;
            }
            if (RAD_SABADO.Checked)
            {
                daysid = 5;
            }
            if (RAD_DOMINGO.Checked)
            {
                daysid = 6;
            }
            LoadDetailsShift(shiftid, daysid);
        }
        private void Btn_ejecutar_Click(object sender, EventArgs e)
        {
            
        }
        private void ProcessDataFilter() 
        {
            var Firsts = from q in registros
                         group q by new { q.UserID } into groups
                         select groups.First();


            var Lasts = from q in registros
                        group q by new { q.UserID } into groups
                        select groups.Last();

            //var Firsts = (from q in registros
            //              group q by new { q.UserID } into grp
            //              select new Registro
            //              {
            //                  UserID = grp.Key.UserID,
            //                  UserName = grp.ElementAt(0).UserName,
            //                  FechaHora_Marca = grp.Min(t => t.FechaHora_Marca),
            //                  Device = grp.FirstOrDefault().Device,
            //                  NameDevice = grp.FirstOrDefault().NameDevice,
            //                  Reference = grp.FirstOrDefault().Reference,
            //                  DateRegistro = grp.FirstOrDefault().DateRegistro,
            //                  HourRegistro = grp.FirstOrDefault().HourRegistro,
            //                  HoraMark = grp.Min(t => t.FechaHora_Marca).ToString("HH:mm tt"),
            //                  NumberRecord = 1
            //              }).FirstOrDefault();

            //var Lasts = (from q in registros
            //group q by new { q.UserID} into grp
            //select new Registro
            //{
            //    UserID = grp.Key.UserID,
            //    UserName = grp.LastOrDefault().UserName,
            //    FechaHora_Marca = grp.Max(t => t.FechaHora_Marca),
            //    Device = grp.LastOrDefault().Device,
            //    NameDevice = grp.LastOrDefault().NameDevice,
            //    Reference = grp.LastOrDefault().Reference,
            //    DateRegistro = grp.LastOrDefault().DateRegistro,
            //    HourRegistro = grp.LastOrDefault().HourRegistro,
            //    HoraMark = grp.Max(t => t.FechaHora_Marca).ToString("HH:mm tt"),
            //    NumberRecord = 2
            //}).LastOrDefault();

            List < Registro > items = new List<Registro>();
            items = Firsts.Union(Lasts).OrderBy(x => x.UserID).ToList();
            //items = Firsts.Union(Lasts).OrderBy(x => x.FechaHora_Marca.TimeOfDay).ToList();
            GridData.DataSource = items;
            Label_Counts.Text = items.Count().ToString() + " Registros Encontrados.";
            if (chk_print_report.Checked)
            {
                ReporteGeneral(items,"FILTERED DATA REPORT");
            }
            if (chk_Import_Excel.Checked)
            {
                DisplayInExcel(items);
            }
            CreateFileTxtDataProches(items);
        }
        private void ConfigFilterForm() 
        {
            // configurar la consulta por dispositivos.

            string d1 = chk_dispo1.Checked ? txt_par1.Text : "";
            string d2 = chk_dispo2.Checked ? txt_par2.Text : "";
            string d3 = chk_dispo3.Checked ? txt_par3.Text : "";
            string d4 = chk_dispo4.Checked ? txt_par4.Text : "";
            string d5 = chk_dispo5.Checked ? txt_par5.Text : "";
            string d6 = chk_dispo6.Checked ? txt_par6.Text : "";
            string d7 = chk_dispo7.Checked ? txt_par7.Text : "";
            string d8 = chk_dispo8.Checked ? txt_par8.Text : "";
            string separetor = ",";

            filter_devices = (string.IsNullOrEmpty(d1) ? string.Empty : d1 + separetor) +
                             (string.IsNullOrEmpty(d2) ? string.Empty : d2 + separetor) +
                             (string.IsNullOrEmpty(d3) ? string.Empty : d3 + separetor) +
                             (string.IsNullOrEmpty(d4) ? string.Empty : d4 + separetor) +
                             (string.IsNullOrEmpty(d5) ? string.Empty : d5 + separetor) +
                             (string.IsNullOrEmpty(d6) ? string.Empty : d6 + separetor) +
                             (string.IsNullOrEmpty(d7) ? string.Empty : d7 + separetor) +
                             (string.IsNullOrEmpty(d8) ? string.Empty : d8);
        }
        private void Conn_AccessDB() 
        {
            OleDbConnection conn = new OleDbConnection
            {
                ConnectionString = Properties.Resources.CONN_STRING_ACCESS
            };

            OleDbCommand comando = new OleDbCommand
            {
                Connection = conn,
                CommandType = CommandType.Text,
                CommandText = "SELECT Record.IdUser,User.Name,Record.RecordTime,Record.MachineNumber," +
                "User.ExternalReference,User.IdentificationNumber,Device.Description " +
                "FROM ([Record] " +
                "LEFT JOIN [User] ON Record.IdUser = User.IdUser) " +
                "LEFT JOIN Device ON Record.MachineNumber = Device.MachineNumber " +
                "WHERE (Record.RecordTime >= @p1 AND  Record.RecordTime <= @p2) AND (Record.MachineNumber=@p3 OR " +
                "Record.MachineNumber=@p4 OR Record.MachineNumber=@p5 OR Record.MachineNumber=@p6 OR " +
                "Record.MachineNumber=@p7 OR Record.MachineNumber=@p8 OR Record.MachineNumber=@p9 OR " +
                "Record.MachineNumber=@p10) AND (Record.IdUser= IIF(@p11 = '',Record.IdUser,@p11))"
            };
            _ = new OleDbParameter("@p1", txt_fecha_desde.Value.Date + txt_hour_desde.Value.TimeOfDay);
            _ = new OleDbParameter("@p2", txt_fecha_hasta.Value.Date + txt_hour_hasta.Value.TimeOfDay);
            _ = new OleDbParameter("@p3", chk_dispo1.Checked ? Convert.ToInt32(txt_par1.Text) : 0);
            _ = new OleDbParameter("@p4", chk_dispo2.Checked ? Convert.ToInt32(txt_par2.Text) : 0);
            _ = new OleDbParameter("@p5", chk_dispo3.Checked ? Convert.ToInt32(txt_par3.Text) : 0);
            _ = new OleDbParameter("@p6", chk_dispo4.Checked ? Convert.ToInt32(txt_par4.Text) : 0);
            _ = new OleDbParameter("@p7", chk_dispo5.Checked ? Convert.ToInt32(txt_par5.Text) : 0);
            _ = new OleDbParameter("@p8", chk_dispo6.Checked ? Convert.ToInt32(txt_par6.Text) : 0);
            _ = new OleDbParameter("@p9", chk_dispo7.Checked ? Convert.ToInt32(txt_par7.Text) : 0);
            _ = new OleDbParameter("@p10", chk_dispo8.Checked ? Convert.ToInt32(txt_par8.Text) : 0);
            _ = new OleDbParameter("@p11", txt_IdUser.Text.Trim());

            comando.Parameters.Add("@p1", OleDbType.Date).Value = txt_fecha_desde.Value.Date + txt_hour_desde.Value.TimeOfDay;
            comando.Parameters.Add("@p2", OleDbType.Date).Value = txt_fecha_hasta.Value.Date + txt_hour_hasta.Value.TimeOfDay;
            comando.Parameters.Add("@p3", OleDbType.VarChar).Value = chk_dispo1.Checked ? Convert.ToInt32(txt_par1.Text) : 0;
            comando.Parameters.Add("@p4", OleDbType.VarChar).Value = chk_dispo2.Checked ? Convert.ToInt32(txt_par2.Text) : 0;
            comando.Parameters.Add("@p5", OleDbType.VarChar).Value = chk_dispo3.Checked ? Convert.ToInt32(txt_par3.Text) : 0;
            comando.Parameters.Add("@p6", OleDbType.VarChar).Value = chk_dispo4.Checked ? Convert.ToInt32(txt_par4.Text) : 0;
            comando.Parameters.Add("@p7", OleDbType.VarChar).Value = chk_dispo5.Checked ? Convert.ToInt32(txt_par5.Text) : 0;
            comando.Parameters.Add("@p8", OleDbType.VarChar).Value = chk_dispo6.Checked ? Convert.ToInt32(txt_par6.Text) : 0;
            comando.Parameters.Add("@p9", OleDbType.VarChar).Value = chk_dispo7.Checked ? Convert.ToInt32(txt_par7.Text) : 0;
            comando.Parameters.Add("@p10", OleDbType.VarChar).Value = chk_dispo8.Checked ? Convert.ToInt32(txt_par8.Text) : 0;
            comando.Parameters.Add("@p11", OleDbType.VarChar).Value = txt_IdUser.Text.Trim();
            conn.Open();
            comando.ExecuteNonQuery();
            _ = new OleDbDataAdapter
            {
                SelectCommand = comando
            };

            OleDbDataReader reader = comando.ExecuteReader();

            // Crear lista de objetos tipo registros.
            registros = new List<Registro>();
            int regNUmbers = 0;
            while (reader.Read())
            {
                regNUmbers += 1;
                Registro ponche = new Registro
                {
                    NumberRecord = regNUmbers,
                    UserID = reader[0].ToString().PadLeft(10, CHARACTER_FILL),
                    UserName = reader[1].ToString(),
                    Device = Convert.ToString(reader.GetInt32(3)).PadRight(10, CHARACTER_FILL),
                    NameDevice = reader[6].ToString(),
                    DateRegistro = Convert.ToString(reader.GetDateTime(2)),
                    FechaHora_Marca = reader.GetDateTime(2),
                    Reference = reader[4].ToString()
                };
                DAYS = reader.GetDateTime(2).Day.ToString().PadLeft(2, '0');
                MONTHS = reader.GetDateTime(2).Month.ToString().PadLeft(2, '0');
                YEARS = reader.GetDateTime(2).Year.ToString().PadLeft(2, '0');
                HOURS = reader.GetDateTime(2).Hour.ToString().PadLeft(2, '0');
                MINUTES = reader.GetDateTime(2).Minute.ToString().PadLeft(2, '0');
                ponche.DateRegistro = DAYS + MONTHS + YEARS;
                ponche.HourRegistro = HOURS + MINUTES;
                registros.Add(ponche);
            }
            conn.Close();
            GridData.DataSource = registros;
            Label_Counts.Text = regNUmbers.ToString()+" Registros Encontrados.";
           

                //CREAR EL ARCHIVO TXT.
                try
                {
                    using (StreamWriter sw = File.CreateText(Properties.Resources.PATH_FILE_TXT + Properties.Resources.FILENAME))
                    {
                        foreach (Registro item in registros)
                        {
                            sw.WriteLine(item.Device + CHAR_SEPARATOR + item.UserID + CHAR_SEPARATOR + item.DateRegistro + CHAR_SEPARATOR + item.HourRegistro);
                        }
                    }
                    // MUESTRA EL TXT DE TEXTBOX
                    using (StreamReader file = new StreamReader(Properties.Resources.PATH_FILE_TXT + Properties.Resources.FILENAME))
                    {
                        var temp = file.ReadToEnd();
                        TXT_TXTARCHIVO.Text = temp;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Error en la ruta donde se guardará el txt.");
                }

            if (chk_Import_Excel.Checked == true)
            {
                DisplayInExcel(registros);
            }
        }
        private void LoadParameters() 
        {
            // Traer los datos de los recursos de la aplicacion
            txt_filename.Text = Properties.Resources.FILENAME;
            Txt_server.Text = Properties.Resources.SERVER_NAME;
            txt_database.Text = Properties.Resources.DATABASE_NAME;
            txt_username.Text = Properties.Resources.USER_SQL;
            txt_password_sql.Text = Properties.Resources.PASSWORD_SQL;
            txt_port_number.Text = Properties.Resources.PORT_COMM;
            TXT_RANGEMINUTES_SALIDAS.Text = Properties.Resources.RANGLE_MIN_SALIDAS;
            FechaHora1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            FechaHora2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 23, 59, 59);
            txt_file_path.Text = Properties.Resources.PATH_FILE_TXT;
            txt_fecha_desde.Value = FechaHora1;
            txt_fecha_hasta.Value = FechaHora2;
            txt_hour_desde.Value = FechaHora1;
            txt_hour_hasta.Value = FechaHora2;
            //Carga manual de parametros.
            txt_par1.Text = "1";
            txt_par2.Text = "2";
            txt_par3.Text = "3";
            txt_par4.Text = "4";
            txt_par5.Text = "5";
            txt_par6.Text = "6";
            txt_par7.Text = "7";
            txt_par8.Text = "8";

        }
    }
    class RecordDistinctUserDevice : IEqualityComparer<Registro>
    {
        // Products are equal if their names and product numbers are equal.
        public bool Equals(Registro x, Registro y)
        {

            //Check whether the compared objects reference the same data.
            if (Object.ReferenceEquals(x, y)) return true;

            //Check whether any of the compared objects is null.
            if (x is null || y is null)
                return false;

            //Check whether the device properties are equal.
            return x.UserID == y.UserID && x.Device == y.Device;
        }

        // If Equals() returns true for a pair of objects
        // then GetHashCode() must return the same value for these objects.

        public int GetHashCode(Registro registro)
        {
            //Check whether the object is null
            if (registro is null) return 0;

            //Get hash code for the Name field if it is not null.
            int hashDeviceName = registro.NameDevice == null ? 0 : registro.NameDevice.GetHashCode();

            //Get hash code for the Code field.
            int hashDeviceCode = registro.Device.GetHashCode();

            //Calculate the hash code for the product.
            return hashDeviceName ^ hashDeviceCode;
        }
    }


}
