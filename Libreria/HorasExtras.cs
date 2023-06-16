using Microsoft.Reporting.WinForms;
using Microsoft.ReportingServices.ReportProcessing.ReportObjectModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI;
using System.Windows.Forms;

namespace WestRockDataPonchesPRO.Libreria
{
    public class HorasExtras
    {
        public SqlConnection micomm;
        public string PathReports = @"C:\Users\NELSON\Desktop\WestRock\WestRockDataPonchesPRO\Reports\";
        public HorasExtras()
        {
            micomm = new SqlConnection();
            micomm.ConnectionString = @"Server=DESKTOP-SOA16M2\SQLEXPRESS;Database=BDBioAdminSQL;Trusted_Connection=True";

        }
        public int ObtenerHorarioEmpleado(int userid, DateTime FechaMarcaje) 
        {
            SqlCommand comando = new SqlCommand();
            comando.Connection = micomm;
            comando.CommandType = CommandType.Text;
            comando.CommandText = "SELECT ShiftId FROM [BDBioAdminSQL].[dbo].[UserShift] WHERE IdUser=@p1 AND @p2 BETWEEN BeginDate AND EndDate";
            SqlParameter p1 = new SqlParameter("@p1", userid);
            SqlParameter p2 = new SqlParameter("@p2", FechaMarcaje);
            comando.Parameters.Add(p1);
            comando.Parameters.Add(p2);
            micomm.Open();
            int result = Convert.ToInt16(comando.ExecuteScalar());
            micomm.Close();
            return result;
        }
        public HORARIO_FIJO ObtenerParametrosHorarios(int shiftid, int dayid) 
        {
            //parametros de horario fijo.
            HORARIO_FIJO hfp = new HORARIO_FIJO();
            //comando sql
            SqlCommand comando = new SqlCommand();
            comando.Connection = micomm;
            comando.CommandType = CommandType.Text;
            comando.CommandText = "SELECT b.IdUser,c.Name,a.ShiftId,d.Description as HORARIO_NAME,DayId,a.description,type,t2inhour,t2outhour,"+
            "t2overtime1beginhour,t2overtime1endhour,t2overtime1factor,t2overtime2beginhour,t2overtime2endhour,"+
            "t2overtime2factor,t2overtime3beginhour,t2overtime3endhour,t2overtime3factor,t2overtime4beginhour," +
            "t2overtime4endhour,t2overtime4factor,t2overtime5beginhour,t2overtime5endhour,t2overtime5factor " +
            "FROM [BDBioAdminSQL].[dbo].[ShiftDetail] a left join [BDBioAdminSQL].[dbo].[UserShift] b " +
            "on a.ShiftId = b.ShiftId left join [BDBioAdminSQL].[dbo].[User] c on c.IdUser = b.IdUser " +
            "LEFT JOIN [BDBioAdminSQL].[dbo].[Shift] d ON d.ShiftId = a.ShiftId " + 
            "where a.ShiftId=@p1 and DayId=@p2 order by DayId";
            SqlParameter p1 = new SqlParameter("@p1", shiftid);
            SqlParameter p2 = new SqlParameter("@p2", dayid);
            comando.Parameters.Add(p1);
            comando.Parameters.Add(p2);
            micomm.Open();
            comando.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter();
            DataTable dt = new DataTable();
            da.SelectCommand = comando;
            da.Fill(dt);
            hfp.Horario_Name = dt.Rows[0]["horario_name"].ToString();
            hfp.Jornada_Start = dt.Rows[0]["t2inhour"].ToString();
            hfp.Jornada_End = dt.Rows[0]["t2outhour"].ToString();
            hfp.Cal1_Start = dt.Rows[0]["t2overtime1beginhour"].ToString();
            hfp.Cal1_End = dt.Rows[0]["t2overtime1endhour"].ToString();
            hfp.Cal1_Factor = dt.Rows[0]["t2overtime1factor"].ToString();
            hfp.Cal2_Start = dt.Rows[0]["t2overtime2beginhour"].ToString();
            hfp.Cal2_End = dt.Rows[0]["t2overtime2endhour"].ToString();
            hfp.Cal2_Factor = dt.Rows[0]["t2overtime2factor"].ToString();
            hfp.Cal3_Start = dt.Rows[0]["t2overtime3beginhour"].ToString();
            hfp.Cal3_End = dt.Rows[0]["t2overtime3endhour"].ToString();
            hfp.Cal3_Factor = dt.Rows[0]["t2overtime3factor"].ToString();
            hfp.Cal4_Start = dt.Rows[0]["t2overtime4beginhour"].ToString();
            hfp.Cal4_End = dt.Rows[0]["t2overtime4endhour"].ToString();
            hfp.Cal4_Factor = dt.Rows[0]["t2overtime4factor"].ToString();
            hfp.Cal5_Start = dt.Rows[0]["t2overtime5beginhour"].ToString();
            hfp.Cal5_End = dt.Rows[0]["t2overtime5endhour"].ToString();
            hfp.Cal5_Factor = dt.Rows[0]["t2overtime5factor"].ToString();
            micomm.Close();
            return hfp;
        }
        public double ObtenerSalarioxHora(int userid) 
        {
            double SalarioHora = 0;
            SqlCommand comando = new SqlCommand();
            comando.Connection = micomm;
            comando.CommandType = CommandType.Text;
            comando.CommandText = "SELECT HourSalary FROM [BDBioAdminSQL].[dbo].[User] WHERE IdUser=@p1";
            SqlParameter p1 = new SqlParameter("@p1", userid);
            comando.Parameters.Add(p1);
            micomm.Open();
            SalarioHora = Convert.ToDouble(comando.ExecuteScalar());
            micomm.Close();
            return SalarioHora;
        }
        public DataTable GetEmpleados() 
        {
            DataTable dt = new DataTable();
            //comando sql para traer los empleados del bioadmin.
            dt.Clear();
            SqlCommand comando = new SqlCommand();
            comando.Connection = micomm;
            comando.CommandType = CommandType.Text;
            comando.CommandText = comando.CommandText = "SELECT a.[IdUser],b.[UserShiftId],b.[ShiftId],c.[Description]," +
            "a.[IdentificationNumber]," +
            "a.[Name],a.[Gender],a.[Title],[Birthday],[PhoneNumber],[MobileNumber],[Address]" +
            ",[ExternalReference],a.[IdDepartment],d.Description as Departamento,[Position],[Active],[Picture],[PictureOrientation]" +
            ",[Privilege],[HourSalary],[Password],[PreferredIdLanguage],[Email],a.[Comment],[ProximityCard]" +
            ",[LastRecord],[LastLogin],[CreatedBy],[CreatedDatetime],[ModifiedBy],[ModifiedDatetime]" +
            ",[AdministratorType],[IdProfile],[DevPassword],[UseShift],[SendSMS],[SMSPhone],[TemplateCode]" +
            ",[ApplyExceptionPermition],[ExceptionPermitionBegin],[ExceptionPermitionEnd] FROM [BDBioAdminSQL].[dbo].[User] a " +
            "LEFT JOIN [BDBioAdminSQL].[dbo].[UserShift] b ON a.IdUser = b.IdUser " +
            "LEFT JOIN [BDBioAdminSQL].[dbo].[Shift] c ON b.ShiftId = c.ShiftId " +
            "LEFT JOIN [BDBioAdminSQL].[dbo].[Department] d ON a.[IdDepartment] = d.[IdDepartment]";
            micomm.Open();
            comando.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = comando;
            da.Fill(dt);
            return dt;
        }
        public bool Repo_HorasExtras_Detalle(string desde,string hasta,string time1, string time2, DataTable dthorasextras) 
        {
            //Definir los parametros del reporte de horas extras.
            ReportParameter[] rParams = new ReportParameter[4]
            {
                new ReportParameter("tdesde", desde),
                new ReportParameter("thasta", hasta),
                new ReportParameter("time1", time1),
                new ReportParameter("time2", time2),
            };
            ReportDataSource rds = new ReportDataSource("horas_extras", dthorasextras);

            //crear la instancia del vbisualizador de reportes.
            ReporteView repoview = new ReporteView();
            repoview.reportViewer1.Reset();
            repoview.reportViewer1.ProcessingMode = ProcessingMode.Local;
            repoview.reportViewer1.LocalReport.ReportPath = @"C:\Users\NELSON\Desktop\WestRock\WestRockDataPonchesPRO\Reports\Reporte_Horas_Extras.rdlc";
            repoview.reportViewer1.LocalReport.DataSources.Clear();
            repoview.reportViewer1.LocalReport.DataSources.Add(rds);
            repoview.reportViewer1.LocalReport.SetParameters(rParams);
            repoview.reportViewer1.LocalReport.Refresh();
            repoview.reportViewer1.RefreshReport();
            repoview.Show();
            return true;
        }

        public bool Repo_Horas_Jornada(string desde, string hasta, string time1, string time2, DataTable dtjornadas) 
        {
            //Definir los parametros del reporte de horas extras.
            ReportParameter[] rParams = new ReportParameter[4]
            {
                new ReportParameter("tdesde", desde),
                new ReportParameter("thasta", hasta),
                new ReportParameter("time1", time1),
                new ReportParameter("time2", time2),
            };
            //creacion de la fuente de datos.
            ReportDataSource rds = new ReportDataSource("Jornadas", dtjornadas);
            //crear la instancia del visualizador de reportes.
            ReporteView repoview = new ReporteView();
            repoview.reportViewer1.Reset();
            repoview.reportViewer1.ProcessingMode = ProcessingMode.Local;
            repoview.reportViewer1.LocalReport.ReportPath = PathReports + "JornadasGeneral.rdlc";
            repoview.reportViewer1.LocalReport.DataSources.Clear();
            repoview.reportViewer1.LocalReport.DataSources.Add(rds);
            repoview.reportViewer1.LocalReport.SetParameters(rParams);
            repoview.reportViewer1.LocalReport.Refresh();
            repoview.reportViewer1.RefreshReport();
            repoview.Show();
            return true;
        }
        
        
        public string GetDepartamento(int userid) 
        {
            SqlCommand comando = new SqlCommand();
            comando.Connection = micomm;
            comando.CommandType = CommandType.Text;
            
            comando.CommandText = 
            "SELECT b.description FROM [BDBioAdminSQL].[dbo].[User] a " +
            "LEFT JOIN [BDBioAdminSQL].[dbo].[Department] b ON a.IdDepartment = b.IdDepartment " +     
            "WHERE a.IdUser=@p1";
            SqlParameter p1 = new SqlParameter("@p1", userid);
            comando.Parameters.Add(p1);
            micomm.Open();
            string result = Convert.ToString(comando.ExecuteScalar());
            micomm.Close();
            return result;
        }
    }
}



