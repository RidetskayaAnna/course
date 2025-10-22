using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Agent.Form
{
    public partial class Treaty : UserControl
    {
        datebase db = new datebase();
        public SqlConnection sqlConnection = null;
        Insurerak insurerak;
        public Treaty(Insurerak insurerak1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            insurerak = insurerak1;
        }
        Form1 form1;
        public Treaty(Form1 form)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            form1 = form;
        }
        Startcs startcs;
        public Treaty(Startcs startcs1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            startcs = startcs1;
        }
        Policyholderak policyholder;
        public Treaty(Policyholderak policyholder1)
        {
            InitializeComponent();
            sqlConnection = new SqlConnection(db.connection);
            policyholder = policyholder1;
        }
        int id2 = 0;
        int id3 = 0;
     
        decimal cof3 = 1;
        public void Tread_load()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select Distinct treaty.idtreaty as [Номер договора],treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик],
treaty.idpolicyholder,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь,
treaty.idvida,vid.name as [Вид страхования],
dateconclusion as [Дата заключения], term as [Количество взносов], 
datestart as [Дата начала действия договора], 
datefinish as [Дата окончания действия договора],
suminsured as [Страховая сумма], 
(select Distinct
Cast(ROUND((
((CASE WHEN(treaty.term='Единовременно' )
THEN  
((treaty.vznos/treaty.suminsured)*100)
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))

else (CASE WHEN(treaty.term='В два срока')
THEN  
((treaty.vznos*2)/treaty.suminsured)*100
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
((treaty.vznos*4)/treaty.suminsured)*100
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))

else null end

) end
) end))
),2)as decimal(18,2)) 
)
as [Корректировочный коэффициент],

(CASE WHEN(vid.name='Медицинское страхование')
THEN Cast(ROUND((0.52),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN Cast(ROUND((1.1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN Cast(ROUND((0.5),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование жизни')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN Cast(ROUND((0.2),2)as decimal(18,2))  
else 1 end
) end
) end
) end
) end
) end
)as [Базовый коэффициент],

(select Distinct
Cast(ROUND((
((CASE WHEN(treaty.term='Единовременно' )
THEN  
((treaty.vznos/treaty.suminsured)*100)

else (CASE WHEN(treaty.term='В два срока')
THEN  
((treaty.vznos*2)/treaty.suminsured)*100

else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
((treaty.vznos*4)/treaty.suminsured)*100
else null end

) end
) end))
),2)as decimal(18,2)) )
as [Страховой тариф],

(select Distinct

((CASE WHEN(treaty.term='Единовременно' )
THEN  
(treaty.vznos)

else (CASE WHEN(treaty.term='В два срока')
THEN  
(treaty.vznos*2)

else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
(treaty.vznos*4)
else null end

) end
) end)))as [Страховая премия],

dateconclusion as [Дата первого взноса],
vznos as [Первый взнос],

(CASE WHEN(treaty.term='В два срока')
THEN  DATEADD(MONTH,6, datestart)
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,3, datestart)
else null end) end) as [Последний день второго взноса],


(CASE WHEN(treaty.term='В два срока' or treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Второй взноса],


(CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,6, datestart)
else null end
)as [Последний день третьего взноса],
(CASE WHEN(treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Третий взноса],

(CASE WHEN(treaty.term='Ежеквартально')
THEN   DATEADD(MONTH,9, datestart)
else null end)as [Последний день четвертого взноса],

(CASE WHEN(treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Четвертый взноса],

treaty.idbid
from
treaty inner join insurer on treaty.idinsurer=insurer.idinsurer
inner join policyholder as policyholder on policyholder.idpolicyholder=treaty.idpolicyholder
inner join vid on vid.idvida=treaty.idvida
inner join bid on bid.idbid=treaty.idbid
inner join pay on treaty.idtreaty=pay.idtreaty
left join correctionfactor on correctionfactor.idvida=vid.idvida
group by
treaty.idtreaty,treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname),
treaty.idpolicyholder,treaty.idvida,vid.name,treaty.vznos,treaty.idbid,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname),
dateconclusion , term , 
datestart , datefinish ,suminsured ,pay.vidpay,
correctionfactor.coefficient,pay.datepay", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.Columns[24].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();

        }
        public void Tread_load3()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select Distinct treaty.idtreaty as [Номер договора],treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик],
treaty.idpolicyholder,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь,
treaty.idvida,vid.name as [Вид страхования],
dateconclusion as [Дата заключения], term as [Количество взносов], 
datestart as [Дата начала действия договора], 
datefinish as [Дата окончания действия договора],
suminsured as [Страховая сумма], 
(select Distinct
Cast(ROUND((
((CASE WHEN(treaty.term='Единовременно' )
THEN  
((treaty.vznos/treaty.suminsured)*100)
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))

else (CASE WHEN(treaty.term='В два срока')
THEN  
((treaty.vznos*2)/treaty.suminsured)*100
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
((treaty.vznos*4)/treaty.suminsured)*100
/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5
else (CASE WHEN(vid.name='Страхование жизни')
THEN 1
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2
else 1 end
) end
) end
) end
) end
) end
))

else null end

) end
) end))
),2)as decimal(18,2)) 
)
as [Корректировочный коэффициент],

(CASE WHEN(vid.name='Медицинское страхование')
THEN Cast(ROUND((0.52),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN Cast(ROUND((1.1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN Cast(ROUND((0.5),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование жизни')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN Cast(ROUND((0.2),2)as decimal(18,2))  
else 1 end
) end
) end
) end
) end
) end
)as [Базовый коэффициент],

(select Distinct
Cast(ROUND((
((CASE WHEN(treaty.term='Единовременно' )
THEN  
((treaty.vznos/treaty.suminsured)*100)

else (CASE WHEN(treaty.term='В два срока')
THEN  
((treaty.vznos*2)/treaty.suminsured)*100

else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
((treaty.vznos*4)/treaty.suminsured)*100
else null end

) end
) end))
),2)as decimal(18,2)) )
as [Страховой тариф],

(select Distinct

((CASE WHEN(treaty.term='Единовременно' )
THEN  
(treaty.vznos)

else (CASE WHEN(treaty.term='В два срока')
THEN  
(treaty.vznos*2)

else (CASE WHEN(treaty.term='Ежеквартально')
THEN  
(treaty.vznos*4)
else null end

) end
) end)))as [Страховая премия],

dateconclusion as [Дата первого взноса],
vznos as [Первый взнос],

(CASE WHEN(treaty.term='В два срока')
THEN  DATEADD(MONTH,6, datestart)
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,3, datestart)
else null end) end) as [Последний день второго взноса],


(CASE WHEN(treaty.term='В два срока' or treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Второй взноса],


(CASE WHEN(treaty.term='Ежеквартально')
THEN  DATEADD(MONTH,6, datestart)
else null end
)as [Последний день третьего взноса],
(CASE WHEN(treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Третий взноса],

(CASE WHEN(treaty.term='Ежеквартально')
THEN   DATEADD(MONTH,9, datestart)
else null end)as [Последний день четвертого взноса],

(CASE WHEN(treaty.term='Ежеквартально')
THEN  vznos
else null end
)as [Четвертый взноса],

treaty.idbid
from
treaty inner join insurer on treaty.idinsurer=insurer.idinsurer
inner join policyholder as policyholder on policyholder.idpolicyholder=treaty.idpolicyholder
inner join vid on vid.idvida=treaty.idvida
inner join bid on bid.idbid=treaty.idbid
inner join pay on treaty.idtreaty=pay.idtreaty
left join correctionfactor on correctionfactor.idvida=vid.idvida
where treaty.idpolicyholder={policyholder.idakk}
group by
treaty.idtreaty,treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname),
treaty.idpolicyholder,treaty.idvida,vid.name,treaty.vznos,treaty.idbid,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname),
dateconclusion , term , 
datestart , datefinish ,suminsured ,pay.vidpay,
correctionfactor.coefficient,pay.datepay", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.Columns[24].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();

        }
        public void Tread_load2()
        {
            sqlConnection.Close();
            sqlConnection.Open();
            DataSet dataSet = new DataSet();
            SqlDataAdapter command = new SqlDataAdapter($@"Select Distinct treaty.idtreaty as [Номер договора],treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname) as [Страховщик],
treaty.idpolicyholder,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname) as Страхователь,
treaty.idvida,vid.name as [Вид страхования],
dateconclusion as [Дата заключения], term as [Количество взносов], 
datestart as [Дата начала действия договора], 
datefinish as [Дата окончания действия договора],
suminsured as [Страховая сумма], 
(select Distinct Cast(ROUND((
((CASE WHEN(treaty.term='Единовременно' )
THEN  ((treaty.vznos/treaty.suminsured)*100) /((CASE WHEN(vid.name='Медицинское страхование') THEN 0.52
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1 else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5 else (CASE WHEN(vid.name='Страхование жизни')
THEN 1 else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2 else 1 end) end) end) end) end) end))
else (CASE WHEN(treaty.term='В два срока')
THEN  ((treaty.vznos*2)/treaty.suminsured)*100/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52 else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1 else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5 else (CASE WHEN(vid.name='Страхование жизни')
THEN 1 else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2 else 1 end) end) end) end) end) end))
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  ((treaty.vznos*4)/treaty.suminsured)*100/((CASE WHEN(vid.name='Медицинское страхование')
THEN 0.52 else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN 1.1 else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN 1 else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN 0.5 else (CASE WHEN(vid.name='Страхование жизни')
THEN 1 else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN 0.2 else 1 end) end) end) end) end) end)) else null end) end) end))),2)as decimal(18,2)) )
as [Корректировочный коэффициент],
(CASE WHEN(vid.name='Медицинское страхование')
THEN Cast(ROUND((0.52),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчатного случая')
THEN Cast(ROUND((1.1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Обязательное страхование гражданской ответственности перевозчика перед пассажирами')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование от несчастных случаев пассажиров железного транспорта')
THEN Cast(ROUND((0.5),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование жизни')
THEN Cast(ROUND((1),2)as decimal(18,2))  
else (CASE WHEN(vid.name='Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности')
THEN Cast(ROUND((0.2),2)as decimal(18,2))  
else 1 end) end) end) end) end) end)as [Базовый коэффициент],
(select Distinct Cast(ROUND((((CASE WHEN(treaty.term='Единовременно' )
THEN  ((treaty.vznos/treaty.suminsured)*100)
else (CASE WHEN(treaty.term='В два срока') THEN  ((treaty.vznos*2)/treaty.suminsured)*100
else (CASE WHEN(treaty.term='Ежеквартально')
THEN  ((treaty.vznos*4)/treaty.suminsured)*100
else null end) end) end))),2)as decimal(18,2)) )as [Страховой тариф],
(select Distinct ((CASE WHEN(treaty.term='Единовременно' )
THEN  (treaty.vznos) else (CASE WHEN(treaty.term='В два срока') THEN  
(treaty.vznos*2) else (CASE WHEN(treaty.term='Ежеквартально') THEN  (treaty.vznos*4)
else null end) end) end)))as [Страховая премия],
dateconclusion as [Дата первого взноса],vznos as [Первый взнос],
(CASE WHEN(treaty.term='В два срока') THEN  DATEADD(MONTH,6, datestart)
else (CASE WHEN(treaty.term='Ежеквартально') THEN  DATEADD(MONTH,3, datestart)
else null end) end) as [Последний день второго взноса],
(CASE WHEN(treaty.term='В два срока' or treaty.term='Ежеквартально')
THEN  vznos else null end)as [Второй взноса],
(CASE WHEN(treaty.term='Ежеквартально') THEN  DATEADD(MONTH,6, datestart)
else null end)as [Последний день третьего взноса],
(CASE WHEN(treaty.term='Ежеквартально')
THEN  vznos else null end
)as [Третий взноса],(CASE WHEN(treaty.term='Ежеквартально')
THEN   DATEADD(MONTH,9, datestart) else null end)as [Последний день четвертого взноса],
(CASE WHEN(treaty.term='Ежеквартально') THEN  vznos
else null end)as [Четвертый взноса],treaty.idbid
from
treaty inner join insurer on treaty.idinsurer=insurer.idinsurer
inner join policyholder as policyholder on policyholder.idpolicyholder=treaty.idpolicyholder
inner join vid on vid.idvida=treaty.idvida
inner join bid on bid.idbid=treaty.idbid
inner join pay on treaty.idtreaty=pay.idtreaty
left join correctionfactor on correctionfactor.idvida=vid.idvida
where treaty.idinsurer={insurerak.idakk}
group by
treaty.idtreaty,treaty.idinsurer,(insurer.firstname+' '+insurer.name+' '+insurer.lastname),
treaty.idpolicyholder,treaty.idvida,vid.name,treaty.vznos,treaty.idbid,
(policyholder.firdtname+' '+policyholder.name+' '+policyholder.lastname),
dateconclusion , term , 
datestart , datefinish ,suminsured ,pay.vidpay,
correctionfactor.coefficient,pay.datepay", sqlConnection);
            command.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0].DefaultView;
            sqlConnection.Close();
            dataGridView1.Columns[1].Visible = false;
            dataGridView1.Columns[3].Visible = false;
            dataGridView1.Columns[5].Visible = false;
            dataGridView1.Columns[24].Visible = false;
            dataGridView1.AllowUserToAddRows = false;
            sqlConnection.Close();

        }

        public void comboBoxinsurer()
        {
            sqlConnection.Open();
            string query = "select idinsurer,(firstname+' '+name+' '+lastname) as i from insurer where datelayoffs is null";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox1.DataSource = dataSet.Tables[0];
            comboBox1.DisplayMember = "i";
            comboBox1.ValueMember = "idinsurer";
            comboBox1.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void comboBoxpolicyholder()
        {
            sqlConnection.Open();
            string query = "select Distinct policyholder.idpolicyholder,(firdtname+' '+name+' '+lastname) as p from policyholder inner join bid on bid.idpolicyholder=policyholder.idpolicyholder where status='Ожидание'";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox2.DataSource = dataSet.Tables[0];
            comboBox2.DisplayMember = "p";
            comboBox2.ValueMember = "idpolicyholder";
            comboBox2.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void comboBoxinsurer2()
        {
            sqlConnection.Open();
            string query = "select idinsurer,(firstname+' '+name+' '+lastname) as i from insurer ";
            SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
            DataSet dataSet = new DataSet();
            sqlDbDataAdapter.Fill(dataSet);
            comboBox7.DataSource = dataSet.Tables[0];
            comboBox7.DisplayMember = "i";
            comboBox7.ValueMember = "idinsurer";
            comboBox7.SelectedIndex = -1;
            sqlConnection.Close();
        }
        public void comboBoxpolicyholder2()
        {
            
                sqlConnection.Open();
                string query = "select idpolicyholder,(firdtname+' '+name+' '+lastname) as p from policyholder";
                SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                DataSet dataSet = new DataSet();
                sqlDbDataAdapter.Fill(dataSet);
                comboBox4.DataSource = dataSet.Tables[0];
                comboBox4.DisplayMember = "p";
                comboBox4.ValueMember = "idpolicyholder";
                comboBox4.SelectedIndex = -1;
                sqlConnection.Close();
            
        }
            
        public void clear()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
           
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker1.MinDate = DateTime.Today;
            dateTimePicker1.MaxDate = DateTime.Today;
            dateTimePicker2.Value = DateTime.Today;
            dateTimePicker2.MinDate = DateTime.Today;
            dateTimePicker2.MaxDate = DateTime.Today.AddMonths(3);
            dateTimePicker3.Value = DateTime.Today.AddDays(365);
            dateTimePicker4.Value = DateTime.Today;
            dateTimePicker5.Value = DateTime.Today;
            dateTimePicker6.Value = DateTime.Today;
            dateTimePicker7.Value = DateTime.Today;
            textBox2.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
        }
        private void Treaty_Load(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView3.RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
                dataGridView3.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                
                    if (insurerak != null)
                    {
                        try
                        {
                            Tread_load2();
                        }
                        catch { }
                    }
                else if (policyholder != null)
                {
                    try
                    {
                        Tread_load3();
                        button4.Visible = false;
                        button3.Visible = false;
                        button5.Visible = false;
                        button7.Visible = false;
                    }
                    catch { }
                }
                else
                    {
                        Tread_load();
                    button4.Visible = true;
                    button3.Visible = true;
                    button5.Visible = true;
                    button7.Visible = true;
                }
              
                comboBoxinsurer();
                comboBoxpolicyholder(); 
            panel3.Visible = false;
            comboBoxinsurer2();
            comboBoxpolicyholder2();
            //// comboBoxobject();
            // panel4.Visible = false;
            // panel2.Visible = false;

           
            string query2 = $@"Select Min(dateconclusion) from treaty";
            System.Data.DataTable data2 = new System.Data.DataTable();
            SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
            command2.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            dateTimePicker8.MinDate = Convert.ToDateTime(row2[column2].ToString());
            dateTimePicker8.MaxDate = DateTime.Today;
            dateTimePicker9.MaxDate = DateTime.Today;
            }
            catch { }
        }

        private void button4_Click(object sender, EventArgs e)
            {
            if (comboBox4.Items.Count != 0)
            {
                if (insurerak !=null)
                {
                    sqlConnection.Open();
                    string query = $@"select idinsurer,(firstname+' '+name+' '+lastname) as i from insurer where datelayoffs is null and idinsurer={insurerak.idakk}";
                    SqlDataAdapter sqlDbDataAdapter = new SqlDataAdapter(query, sqlConnection);
                    DataSet dataSet = new DataSet();
                    sqlDbDataAdapter.Fill(dataSet);
                    comboBox1.DataSource = dataSet.Tables[0];
                    comboBox1.DisplayMember = "i";
                    comboBox1.ValueMember = "idinsurer";
                    comboBox1.SelectedIndex = 0;
                    sqlConnection.Close();
                    //comboBox1.Enabled= false;
                }
                else
                {
                    comboBox1.Enabled = true;
                }

                panel6.Visible = true;
            panel2.Visible = false;
                panel3.Visible = false;
                panel7.Visible = false;
            id2 = 0;
            id3 = 0;
            clear();
        }else
            {
                MessageBox.Show("Заявок нет!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
             }

}

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker3.Value = dateTimePicker2.Value.AddDays(365);
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textBox7.Text != "")
            {
                if (comboBox6.SelectedIndex == 0)
                {
                    textBox8.Text = textBox7.Text;
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";

                }
                else if (comboBox6.SelectedIndex == 1)
                {
                    textBox8.Text = Convert.ToString(Math.Round((Convert.ToDecimal(textBox7.Text) / 2),2));
                    textBox9.Text = textBox8.Text;
                    textBox10.Text = "";
                    textBox11.Text = "";
                }
                else if (comboBox6.SelectedIndex == 2)
                {
                    textBox8.Text = Convert.ToString(Math.Round((Convert.ToDecimal(textBox7.Text) / 4), 2));
                    textBox9.Text = textBox8.Text;
                    textBox10.Text = textBox8.Text;
                    textBox11.Text = textBox8.Text;
                }
            }
            if (comboBox6.SelectedIndex == 0)
            {
                dateTimePicker4.Value = dateTimePicker1.Value;
                dateTimePicker5.Value = DateTime.Today;
                dateTimePicker6.Value = DateTime.Today;
                dateTimePicker7.Value = DateTime.Today;

            }
            else if (comboBox6.SelectedIndex == 1)
            {
                dateTimePicker4.Value = dateTimePicker1.Value;
                dateTimePicker7.Value = dateTimePicker2.Value.AddMonths(6);
                dateTimePicker5.Value = DateTime.Today;
                dateTimePicker6.Value = DateTime.Today;
            }
            else if (comboBox6.SelectedIndex == 2)
            {
                dateTimePicker4.Value = dateTimePicker1.Value;
                dateTimePicker7.Value = dateTimePicker2.Value.AddMonths(3);
                dateTimePicker6.Value = dateTimePicker2.Value.AddMonths(6);
                dateTimePicker5.Value = dateTimePicker2.Value.AddMonths(9);
            }
        }

        Decimal cof1 = 1;
        String cof = "";
        int kk = 0;
        int znak = 0;
        int countt = 0;
        int y = 0;
        int idc = 0;
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            idc = 0;
            cof1 = 1;
            try
            {
                string query2 = $@"Select Count(coefficient) from correctionfactor inner join vid on vid.idvida=correctionfactor.idvida where correctionfactor.name='Страховая сумма' and vid.idvida={id3}";
                System.Data.DataTable data2 = new System.Data.DataTable();
                SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                command2.Fill(data2);
                DataColumn column2 = data2.Columns[0];
                DataRow row2 = data2.Rows[0];
                kk = Convert.ToInt32(row2[column2].ToString());

                string query1 = $@"Select coefficient,correctionfactor.note,idcorrectionfactor 
from correctionfactor inner join vid on vid.idvida=correctionfactor.idvida 
where correctionfactor.name='Страховая сумма' and vid.idvida={id3}";
                System.Data.DataTable data = new System.Data.DataTable();
                SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                command1.Fill(data);
                for (int i = 0; i < kk; i++)
                {
                    DataColumn column = data.Columns[1];
                    DataRow row = data.Rows[i];
                    cof = (row[column].ToString());
                    char[] o = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
                    znak = cof.LastIndexOfAny(o);
                    countt = cof.IndexOfAny(o);
                    y = Convert.ToInt32(cof.Substring(countt, znak - countt + 1));

                    if (countt == 2)
                    {
                        if (cof.Substring(0, 2) == ">=")
                        {
                            if (Convert.ToDecimal(textBox2.Text) >= y)
                            {
                                DataColumn column3 = data.Columns[0];
                                DataRow row3 = data.Rows[i];
                                cof1 = Convert.ToDecimal(row3[column3].ToString());
                                DataColumn column33 = data.Columns[2];
                                DataRow row33 = data.Rows[i];
                                idc = Convert.ToInt32(row33[column33].ToString());
                                break;
                            }
                        }
                        else if (cof.Substring(0, 2) == "<=")
                        {
                            if (Convert.ToDecimal(textBox2.Text) <= y)
                            {
                                DataColumn column4 = data.Columns[0];
                                DataRow row4 = data.Rows[i];
                                cof1 = Convert.ToDecimal(row4[column4].ToString());
                                DataColumn column33 = data.Columns[2];
                                DataRow row33 = data.Rows[i];
                                idc = Convert.ToInt32(row33[column33].ToString());
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (cof.Substring(0, 1) == ">")
                        {
                            if (Convert.ToDecimal(textBox2.Text) > y)
                            {
                                DataColumn column4 = data.Columns[0];
                                DataRow row4 = data.Rows[i];
                                cof1 = Convert.ToDecimal(row4[column4].ToString());
                                DataColumn column33 = data.Columns[2];
                                DataRow row33 = data.Rows[i];
                                idc = Convert.ToInt32(row33[column33].ToString());
                                break;
                            }
                        }
                        else if (cof.Substring(0, 1) == "=")
                        {
                            if (Convert.ToDecimal(textBox2.Text) == y)
                            {
                                DataColumn column4 = data.Columns[0];
                                DataRow row4 = data.Rows[i];
                                cof1 = Convert.ToDecimal(row4[column4].ToString());
                                DataColumn column33 = data.Columns[2];
                                DataRow row33 = data.Rows[i];
                                idc = Convert.ToInt32(row33[column33].ToString());
                                break;
                            }
                        }
                        else
                        {
                            string query3 = $@"Select coefficient,correctionfactor.note,idcorrectionfactor 
from correctionfactor inner join vid on vid.idvida=correctionfactor.idvida 
where correctionfactor.name='Страховая сумма' and vid.idvida={id3} order by correctionfactor.note ASC";
                            System.Data.DataTable data3 = new System.Data.DataTable();
                            SqlDataAdapter command3 = new SqlDataAdapter(query3, sqlConnection);
                            command3.Fill(data3);
                            for (int j = 0; j < kk; j++)
                            {
                                DataColumn column1 = data3.Columns[1];
                                DataRow row1 = data3.Rows[j];
                                cof = (row1[column1].ToString());
                                char[] p = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' };
                                znak = cof.LastIndexOfAny(p);
                                countt = cof.IndexOfAny(p);
                                y = Convert.ToInt32(cof.Substring(countt, znak - countt + 1));

                                if (cof.Substring(0, 1) == "<")
                                {
                                    if (Convert.ToDecimal(textBox2.Text) < y)
                                    {
                                        DataColumn column3 = data3.Columns[0];
                                        DataRow row3 = data3.Rows[j];
                                        cof1 = Convert.ToDecimal(row3[column3].ToString());
                                        DataColumn column33 = data3.Columns[2];
                                        DataRow row33 = data3.Rows[j];
                                        idc = Convert.ToInt32(row33[column33].ToString());
                                        break;

                                    }
                                }

                            }
                        }
                    }

                }


            }
            catch { }
            textBox5.Text = Convert.ToString(Math.Round(cof1 * cof2, 2));
            if (textBox6.Text != "" && textBox2.Text != "")
            {
                textBox7.Text = Convert.ToString(Math.Round(((Convert.ToDecimal(textBox2.Text) * (Math.Round((cof1 * cof2 * cof3), 2))) / 100),2));
            }
            else if(textBox2.Text==""){
                textBox7.Text = "";
                textBox8.Text = "";
                textBox9.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
            }
        }

        private void comboBox5_TextChanged(object sender, EventArgs e)
        {

        }
        Decimal cof2 = 1;
        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox5.SelectedIndex == 0)
            {
                cof2 = 1;
            }
            else if (comboBox5.SelectedIndex == 1)
            {
                cof2 = Convert.ToDecimal(0.9);
            }
            textBox5.Text = Convert.ToString(Math.Round(cof1 * cof2, 2));
        }
        int kkk = 0;
      //  Decimal cof3 = 1;
        String mat = "";
        int idobi = 0;
        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void comboBox3_Leave(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                textBox6.Text = Convert.ToString(Math.Round((cof1 * cof2 * cof3), 2));
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            if (textBox4.Text != "" && textBox5.Text != "")
            {
                textBox6.Text = Convert.ToString(Math.Round((cof1 * cof2 * cof3), 2));
            }
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            if (textBox6.Text != "" && textBox2.Text != "")
            {
                textBox7.Text = Convert.ToString(Math.Round(((Convert.ToDecimal(textBox2.Text) * (Math.Round((cof1 * cof2 * cof3), 2))) / 100), 2));
            }
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }
    
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            if (textBox7.Text != "")
            {
                if (comboBox6.SelectedIndex == 0)
                {
                    textBox8.Text = textBox7.Text;
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";

                }
             if (comboBox6.SelectedIndex == 1)
                {
                    textBox8.Text = Convert.ToString((Convert.ToDecimal(textBox7.Text) / 2));
                    textBox9.Text = textBox8.Text;
                    textBox10.Text = "";
                    textBox11.Text = "";
                }
                else if (comboBox6.SelectedIndex == 2)
                {
                    textBox8.Text = Convert.ToString((Convert.ToDecimal(textBox7.Text) / 4));
                    textBox9.Text = textBox8.Text;
                    textBox10.Text = textBox8.Text;
                    textBox11.Text = textBox8.Text;
                }
            }
        }
        int idtre = 0;
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {


                if (comboBox6.SelectedIndex != -1 && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1 && comboBox5.SelectedIndex != -1 && textBox2.Text != "")
                {
                    if (textBox8.Text != "")
                    {
                        sqlConnection.Open();
                        SqlCommand command = new SqlCommand($@"INSERT INTO [treaty](idinsurer,idpolicyholder,idvida,
                            dateconclusion,term,suminsured,datestart,datefinish,vznos,idbid)
                           VALUES (@in,@po,@iv,@date,@t,@s,@ds,@df,@vz,@ib);", sqlConnection);
                        command.Parameters.AddWithValue("@in", (comboBox1.SelectedValue));
                        command.Parameters.AddWithValue("@po", (comboBox2.SelectedValue));
                        command.Parameters.AddWithValue("@iv", (id3));
                        command.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                        command.Parameters.AddWithValue("@t", (comboBox6.SelectedItem));
                        command.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox2.Text));
                        command.Parameters.AddWithValue("@ds", (dateTimePicker2.Value));
                        command.Parameters.AddWithValue("@df", (dateTimePicker3.Value));
                        command.Parameters.AddWithValue("@vz", Convert.ToDecimal(textBox8.Text));
                        command.Parameters.AddWithValue("@ib", (id2));
                        command.ExecuteNonQuery();

                        sqlConnection.Close();
                        sqlConnection.Open();
                        SqlCommand command2 = new SqlCommand($@"UPDATE bid SET status=@s WHERE idbid=@id", sqlConnection);
                        command2.Parameters.AddWithValue("@s", ("Оформлен"));
                        command2.Parameters.AddWithValue("@id", (id2));
                        command2.ExecuteNonQuery();
                        sqlConnection.Close();

                    }
                if (insurerak !=null)
                {
                    Tread_load2();
                }
                else
                {
                    Tread_load();
                }
                panel2.Visible = false;
                    button11.Visible = false;
                    button12.Visible = false;


            }
            else
            {
                MessageBox.Show("Заполните поля!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        
    }
            catch { }
        }

        private void button11_Leave(object sender, EventArgs e)
        {

        }

        private void button11_MouseUp(object sender, MouseEventArgs e)
        {
            sqlConnection.Close();
            string query2 = $@"Select max(idtreaty) from treaty";
            DataTable data2 = new DataTable();
            SqlDataAdapter command3 = new SqlDataAdapter(query2, sqlConnection);
            command3.Fill(data2);
            DataColumn column2 = data2.Columns[0];
            DataRow row2 = data2.Rows[0];
            idtre = Convert.ToInt32(row2[column2].ToString());
            if (comboBox6.SelectedIndex != -1 && comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1 && comboBox5.SelectedIndex != -1 && textBox2.Text != "")
            {
                sqlConnection.Open();
                SqlCommand command2 = new SqlCommand($@"INSERT INTO [pay](idtreaty,datepay,summa,vidpay) VALUES (@idt,@date,@s,@v);", sqlConnection);
                command2.Parameters.AddWithValue("@idt", (idtre));
                command2.Parameters.AddWithValue("@date", (dateTimePicker1.Value));
                command2.Parameters.AddWithValue("@s", Convert.ToDecimal(textBox8.Text));
                command2.Parameters.AddWithValue("@v", (comboBox5.SelectedItem));
                command2.ExecuteNonQuery();
                sqlConnection.Close();
                clear();
                if (insurerak !=null)
                {
                    Tread_load2();
                }
                else
                {
                    Tread_load();
                }
                panel2.Visible = false;
                
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (j != 1 && j != 3 && j != 4 && j != 6 && j != 8)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Value != null)
                        {
                            if (dataGridView1.Rows[i].Cells[j].Value.ToString().ToLower().Contains(textBox1.Text.ToLower()) && textBox1.Text != "")
                            {
                                dataGridView1.Rows[i].Selected = true;
                                dataGridView1.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                                dataGridView1.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                                break;
                            }
                            else
                            {
                                dataGridView1.Rows[i].Selected = false;
                                dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;


                            }
                        }
                    }
                }
            }
        }

        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker9.MinDate = dateTimePicker8.Value;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ///// 1 по 3 ////

            if (checkBox1.Checked == true && checkBox2.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }

            /////////////////

            //// 3 по 2 ////

            //1
            if (checkBox1.Checked == true && checkBox2.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
           //2
           if (checkBox1.Checked == true && checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
            //3
            if (checkBox3.Checked == true && checkBox2.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue)) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
             ///////////////// 

             //// 3 по 1 ///// 

             //1
             if (checkBox1.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;
                    if ((Convert.ToInt32(dataGridView1[1, i].Value.ToString()) == Convert.ToInt32(comboBox7.SelectedValue)))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
           //2
           if (checkBox2.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToInt32(dataGridView1[3, i].Value.ToString()) == Convert.ToInt32(comboBox4.SelectedValue)))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }

                }

            }
            else
            //3
            if (checkBox3.Checked == true)
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    dataGridView1.CurrentCell = null;
                    dataGridView1.Rows[i].Visible = false;

                    if ((Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) >= Convert.ToDateTime(dateTimePicker8.Value.ToString())) && (Convert.ToDateTime(dataGridView1[7, i].Value.ToString()) <= Convert.ToDateTime(dateTimePicker9.Value.ToString())))
                    {
                        dataGridView1.Rows[i].Visible = true;
                    }
                    else
                    {
                        dataGridView1.Rows[i].Visible = false;
                    }
                }
            }
            else if(checkBox3.Checked == false&& checkBox2.Checked == false&&checkBox1.Checked == false)
            {
                Tread_load();
            }
            /////
        }

        private void button7_Click(object sender, EventArgs e)
        {
          
          //  dataGridView2.Visible = false;
            clear();
            panel2.Visible = false;
            dataGridView1.Enabled = true;
            if (panel3.Visible == true)
            {
                if (insurerak.button2.Text == "Профиль")
                {
                    checkBox1.Enabled = false;
                    comboBox7.Enabled = false;
                    Tread_load2();
                }
                else
                {
                    checkBox1.Enabled = true;
                    comboBox7.Enabled = true;
                    Tread_load();
                }
                panel3.Visible = false;
                
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (48).png");
            }
            else
            {
                if (insurerak!=null)
                {
                    checkBox1.Enabled = false;
                    comboBox7.Enabled = false;
                    
                }
                else
                {
                    checkBox1.Enabled = true;
                    comboBox7.Enabled = true;
                    
                }
                panel3.Visible = true;
                dataGridView1.Height = 407;
                dateTimePicker8.Value = DateTime.Today;
                dateTimePicker9.Value = DateTime.Today;
                comboBox7.SelectedItem = -1;
                comboBox4.SelectedItem = -1;
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
                button7.Image = new Bitmap(@"D:\College\4kurs\Praktica4Kurs\proga\Agent\Agent\Resources\pngwing.com (47).png");
            }
        }



        private void ReplaceWordStubs(string stubToReplace, string text, Word.Document WordDoc)
        {
            var range = WordDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }
        string strax = "";
        string strax2 = "";
        string pas = "";
        string vznos = "";
        private void button3_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = false;
            panel7.Visible = false;
            panel6.Visible = false;
            if ( id != 0){
                
                    if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Медицинское страхование")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\med\med.docx";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);               
                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0,2), document);
                    ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
                    ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);
                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
                    strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();

                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
email as Почта
from policyholder inner join city on policyholder.idcity=city.idcity 
inner join position on position.idposition=policyholder.idwork 
inner join work on work.idwork=position.idwork
 where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    for (int i = 0; i < 10; i++)
                    {
                        DataColumn column = data.Columns[i];
                        DataRow row = data.Rows[0];


                        if (i == 1)
                        {
                            strax = strax + "; " + row[column].ToString();

                            pas = pas + row[column].ToString();
                        }
                        else
                        if (i == 2)
                        {
                            strax = strax + ", г." + row[column].ToString();
                            pas = pas + ", г." + row[column].ToString();
                        }
                        else
                        if (i == 3)
                        {
                            strax = strax + ", " + row[column].ToString();
                            pas = pas + ", " + row[column].ToString();
                        }
                        else
                        if (i == 4)
                        {
                            strax = strax + ";";
                            strax2 = strax2 + " паспорт " + row[column].ToString();
                        }
                        else
                        if (i == 5)
                        {
                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
                        }
                        else
                        if (i == 6)
                        {
                            strax2 = strax2 + ", выдан  " + row[column].ToString();
                        }
                        else
                        if (i == 7)
                        {
                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
                        }
                        else if (i == 8)
                        {
                            pas=pas + ", " + row[column].ToString();
                        }
                        
                    }

                    ReplaceWordStubs("{страхователь}", strax, document);
                    ReplaceWordStubs("{паспорт}", strax2, document);
                    ReplaceWordStubs("{фио}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    DataColumn column5 = data.Columns[0];
                    DataRow row5 = data.Rows[0];
                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{паспорт}", strax2+pas, document);

                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

                    DataColumn column10 = data.Columns[11];
                    DataRow row10 = data.Rows[0];
                    if (row10[column10].ToString() == "Д-1")
                    {
                    ReplaceWordStubs("{х}","X" , document);
                    ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                    }
                    else if(row10[column10].ToString() == "Д-2")
                    {
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "X", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                    }else if (row10[column10].ToString() == "Д-3")
                    {
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "X", document);
                        ReplaceWordStubs("{х}", "", document);
                    }else
                    {
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "X", document);
                    }

                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);

                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);

                    ReplaceWordStubs("{коэф}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

                    string query2 = $@"Select  pay.vidpay,datepay,summa
from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
 where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                    DataTable data2 = new DataTable();
                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                    command2.Fill(data2);
                        DataColumn column1 = data2.Columns[0];
                        DataRow row1 = data2.Rows[0];
                    ReplaceWordStubs("{оплата}", row1[column1].ToString(), document);

                    if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
                    {
                        ReplaceWordStubs("{х}", "X", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        
                    }
                    else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
                    {
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "X", document);
                        ReplaceWordStubs("{х}", "", document);
                      
                    }
                     else
                    {
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "", document);
                        ReplaceWordStubs("{х}", "X", document);
                    }
                    Random random = new Random();
                    ReplaceWordStubs("{страх}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0,10), document);
                    ReplaceWordStubs("{платеж}", row1[column1].ToString()+" №"+ random.Next(0,100), document);

                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString()!="")
                    {
                        vznos = vznos + " " + dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0,10) + "; " + dataGridView1.CurrentRow.Cells[19].Value.ToString() + ";";
                        if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
                        {
                            vznos = vznos + " " + dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0,10) + "; " + dataGridView1.CurrentRow.Cells[21].Value.ToString() + ";";
                            vznos = vznos + " " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0,10) + "; " + dataGridView1.CurrentRow.Cells[23].Value.ToString() + ";";

                        }
                    }
                    else { vznos = ""; }
                    ReplaceWordStubs("{взносы}", vznos, document);
                    ReplaceWordStubs("{страхов}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);

                    DataColumn column11 = data.Columns[8];
                    DataRow row11 = data.Rows[0];

                    ReplaceWordStubs("{тел}", row11[column11].ToString(), document);
                    application.Visible = true;
                
            }
                    else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчатного случая")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\nes\nes.docx";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);
                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
                    
                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
                    strax = dataGridView1.CurrentRow.Cells[4].Value.ToString();

                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
email as Почта,position.harmhul
from policyholder inner join city on policyholder.idcity=city.idcity 
inner join position on position.idposition=policyholder.idwork 
inner join work on work.idwork=position.idwork
 where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    for (int i = 0; i < 10; i++)
                    {
                        DataColumn column = data.Columns[i];
                        DataRow row = data.Rows[0];


                        if (i == 1)
                        {
                            strax = strax + "; " + row[column].ToString();

                            pas = pas + row[column].ToString();
                        }
                        else
                        if (i == 2)
                        {
                            strax = strax + ", г." + row[column].ToString();
                            pas = pas + ", г." + row[column].ToString();
                        }
                        else
                        if (i == 3)
                        {
                            strax = strax + ", " + row[column].ToString();
                            pas = pas + ", " + row[column].ToString();
                        }
                        else
                        if (i == 4)
                        {
                            strax = strax + ";";
                            strax2 = strax2 + " паспорт " + row[column].ToString();
                        }
                        else
                        if (i == 5)
                        {
                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
                        }
                        else
                        if (i == 6)
                        {
                            strax2 = strax2 + ", выдан  " + row[column].ToString();
                        }
                        else
                        if (i == 7)
                        {
                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
                        }
                        else if (i == 8)
                        {
                            pas = pas + ", " + row[column].ToString();
                        }

                    }

                    ReplaceWordStubs("{страхователь}", strax, document);
                    ReplaceWordStubs("{паспорт}", strax2, document);

                    string query45 = $@"select date from treaty 
inner join bid on bid.idbid=treaty.idbid where treaty.idbid={dataGridView1.CurrentRow.Cells[24].Value.ToString()}";
                    DataTable data45 = new DataTable();
                    SqlDataAdapter command45 = new SqlDataAdapter(query45, sqlConnection);
                    command45.Fill(data45);
                    DataColumn column45 = data45.Columns[0];
                    DataRow row45 = data45.Rows[0];

                    ReplaceWordStubs("{датез}", row45[column45].ToString().Substring(0, 10), document);

                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);

                     DataColumn column5 = data.Columns[0];
                    DataRow row5 = data.Rows[0];
                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);

                    DataColumn column15 = data.Columns[9];
                    DataRow row15 = data.Rows[0];
                    DataColumn column16 = data.Columns[10];
                    DataRow row16 = data.Rows[0];
                    ReplaceWordStubs("{работа}", row15[column15].ToString()+"; "+ row16[column16].ToString(), document);
                    
                    DataColumn column10 = data.Columns[11];
                    DataRow row10 = data.Rows[0];
                    if (row10[column10].ToString() == "Д-1")
                    {
                        ReplaceWordStubs("{й}", "", document);
                        ReplaceWordStubs("{ц}", "Х", document);
                        
                    }
                    else if (row10[column10].ToString() == "Д-2")
                    {
                        ReplaceWordStubs("{й}", "", document);
                        ReplaceWordStubs("{ц}", "Х", document);

                    }
                    else if (row10[column10].ToString() == "Д-3")
                    {
                        ReplaceWordStubs("{й}", "", document);
                        ReplaceWordStubs("{ц}", "Х", document);
                    }
                    else
                    {
                        ReplaceWordStubs("{й}", "Х", document);
                        ReplaceWordStubs("{ц}", "", document);
                      
                    }

                    DataColumn column11 = data.Columns[12];
                    DataRow row11 = data.Rows[0];
                    if (row11[column11].ToString() == "Да")
                    {
                        ReplaceWordStubs("{у}", "Х", document);
                        ReplaceWordStubs("{к}", "", document);

                    }
                    else 
                    {
                        ReplaceWordStubs("{у}", "", document);
                        ReplaceWordStubs("{к}", "X", document);

                    }

                    DataColumn column13 = data.Columns[14];
                    DataRow row13 = data.Rows[0];

                    ReplaceWordStubs("{е}", row13[column13].ToString(), document);

                    List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна","Зайцев Владимир Даниилович" };

                    Random rnd = new Random();
                    int randIndex = rnd.Next(numbers.Count);
                    string random = numbers[randIndex];

                    ReplaceWordStubs("{выг}",random, document);
                    ReplaceWordStubs("{выг}", random, document);

                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
                    ReplaceWordStubs("{бст}", dataGridView1.CurrentRow.Cells[13].Value.ToString(), document);
                    ReplaceWordStubs("{кк}", dataGridView1.CurrentRow.Cells[12].Value.ToString(), document);
                    ReplaceWordStubs("{страхт}", dataGridView1.CurrentRow.Cells[14].Value.ToString(), document);
                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

 string query2 = $@"Select  pay.vidpay,datepay,summa
from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
 where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                    DataTable data2 = new DataTable();
                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                    command2.Fill(data2);

                    DataColumn column1 = data2.Columns[0];
                    DataRow row1 = data2.Rows[0];
                    if (row1[column1].ToString() == "Наличные")
                    {
                        ReplaceWordStubs("{н}", "X", document);
                        ReplaceWordStubs("{г}", "", document);
                        ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString(), document);
                        ReplaceWordStubs("{оплата}", "", document);
                    }
                    else
                    {
                        ReplaceWordStubs("{н}", "", document);
                        ReplaceWordStubs("{г}", "Х", document);
                        ReplaceWordStubs("{оплата}", "", document);
                        Random random22 = new Random();
                        ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 10)+ " №" + random22.Next(0, 100), document);
                    }
                  

                    if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "Единовременно")
                    {
                        ReplaceWordStubs("{ш}", "X", document);
                        ReplaceWordStubs("{щ}", "", document);
                        ReplaceWordStubs("{з}", "", document);

                    }
                    else if (dataGridView1.CurrentRow.Cells[8].Value.ToString() == "В два срока")
                    {
                        ReplaceWordStubs("{ш}", "", document);
                        ReplaceWordStubs("{щ}", "X", document);
                        ReplaceWordStubs("{з}", "", document);

                    }
                    else
                    {
                        ReplaceWordStubs("{ш}", "", document);
                        ReplaceWordStubs("{щ}", "", document);
                        ReplaceWordStubs("{з}", "X", document);
                    }

                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);

                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "") {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                        ReplaceWordStubs("{датат}", "", document);
                        ReplaceWordStubs("{датач}", "", document);
                        ReplaceWordStubs("{взнос}", "", document);
                        ReplaceWordStubs("{взнос}", "", document);
                    }
                    else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "") {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                        ReplaceWordStubs("{датат}", dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10), document);
                        ReplaceWordStubs("{датач}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                        ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                    }
                    else
                    {
                        ReplaceWordStubs("{датав}", "", document);
                        ReplaceWordStubs("{взнос}", "", document);
                        ReplaceWordStubs("{датат}", "", document);
                        ReplaceWordStubs("{датач}", "", document);
                        ReplaceWordStubs("{взнос}", "", document);
                        ReplaceWordStubs("{взнос}", "", document);
                    }

                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);

                    ReplaceWordStubs("{время}",DateTime.Now.ToString().Substring(10,6) , document);
                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 2), document);
                    ReplaceWordStubs("{месяц}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(3, 2), document);
                    ReplaceWordStubs("{год}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(6, 4), document);
                   
                    DataColumn column12 = data.Columns[8];
                    DataRow row12 = data.Rows[0];

                    ReplaceWordStubs("{тел}", row12[column12].ToString(), document);
                    application.Visible = true;
                }
                    else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Обязательное страхование гражданской ответственности перевозчика перед пассажирами")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\pere\pere.docx";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);
                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0,10), document);
                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{страхс}", dataGridView1.CurrentRow.Cells[11].Value.ToString(), document);
                       string query2 = $@"Select  pay.vidpay,datepay,summa
from pay inner join  treaty on treaty.idtreaty=pay.idtreaty
 where treaty.idtreaty={dataGridView1.CurrentRow.Cells[0].Value.ToString()}";
                    DataTable data2 = new DataTable();
                    SqlDataAdapter command2 = new SqlDataAdapter(query2, sqlConnection);
                    command2.Fill(data2);
                    DataColumn column1 = data2.Columns[0];
                    DataRow row1 = data2.Rows[0];
                      ReplaceWordStubs("{оплата}", row1[column1].ToString().ToLower(), document);
                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
                    application.Visible = true;
                }
                    else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование от несчастных случаев пассажиров железного транспорта")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\pas\pas.docx";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);
                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);
                    ReplaceWordStubs("{оплата}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);
                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                    ReplaceWordStubs("{число}", dataGridView1.CurrentRow.Cells[16].Value.ToString().Substring(0, 2), document);
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
                    application.Visible = true;
                }
                    else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование жизни")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\live\live.docx";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);
                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);


                    string query1 = $@"Select dateb as [Дата рождения],city.indexcity,
city.name as [Город прописки],address as [Адрес], passport as [Паспорт],
numar as [Идентификационный номер],organ as [Орган, выдавший паспорт],
datep as [Дата выдачи паспорта],phone as Телефон,work.name as [Место работы],
position.name as [Должность], heal as [Группа здоровья],sport as [Занятие спортом],
email as Почта,position.harmhul
from policyholder inner join city on policyholder.idcity=city.idcity 
inner join position on position.idposition=policyholder.idwork 
inner join work on work.idwork=position.idwork
 where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                    for (int i = 0; i < 10; i++)
                    {
                        DataColumn column = data.Columns[i];
                        DataRow row = data.Rows[0];


                        if (i == 1)
                        {
                            strax = strax + " " + row[column].ToString();

                            pas = pas + row[column].ToString();
                        }
                        else
                        if (i == 2)
                        {
                            strax = strax + ", г." + row[column].ToString();
                            pas = pas + ", г." + row[column].ToString();
                        }
                        else
                        if (i == 3)
                        {
                            strax = strax + ", " + row[column].ToString();
                            pas = pas + ", " + row[column].ToString();
                        }
                        else
                        if (i == 4)
                        {
                            strax = strax + ";";
                            strax2 = strax2 + " паспорт " + row[column].ToString();
                        }
                        else
                        if (i == 5)
                        {
                            strax2 = strax2 + ", идентификационный номер " + row[column].ToString();
                        }
                        else
                        if (i == 6)
                        {
                            strax2 = strax2 + ", выдан  " + row[column].ToString();
                        }
                        else
                        if (i == 7)
                        {
                            strax2 = strax2 + "; " + row[column].ToString().Substring(0, 10);
                        }
                        else if (i == 8)
                        {
                            pas = pas + ", " + row[column].ToString();
                        }

                    }

                    ReplaceWordStubs("{паспорт}", strax+strax2, document);

                    DataColumn column5 = data.Columns[0];
                    DataRow row5 = data.Rows[0];
                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датаро}", row5[column5].ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{паспорт}", strax + strax2, document);
                    List<string> numbers = new List<string>() { "Никитин Николай Михайлович", "Иванова Кристина Тимофеевна", "Зайцев Владимир Даниилович" };

                    Random rnd = new Random();
                    int randIndex = rnd.Next(numbers.Count);
                    string random = numbers[randIndex];

                    ReplaceWordStubs("{выг}", random, document);

                    ReplaceWordStubs("{страхс}",Convert.ToString( Convert.ToDecimal( dataGridView1.CurrentRow.Cells[11].Value.ToString())/2), document);
                    ReplaceWordStubs("{страхсс}", Convert.ToString((Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString())/3)), document);
                    ReplaceWordStubs("{страхссс}", Convert.ToString(Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString())- (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 3)- (Convert.ToDecimal(dataGridView1.CurrentRow.Cells[11].Value.ToString()) / 2)), document);
                    ReplaceWordStubs("{страхп}", dataGridView1.CurrentRow.Cells[15].Value.ToString(), document);

                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
                    {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
                       
                    }
                    else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
                    {
                        
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
      
                    }
                    else
                    {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
                       
                    }

                   
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{страховщик}", dataGridView1.CurrentRow.Cells[2].Value.ToString(), document);

                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    application.Visible = true;
                }
                    else if (dataGridView1.CurrentRow.Cells[6].Value.ToString() == "Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности")
                    {
                    strax = "";
                    strax2 = "";
                    pas = "";
                    vznos = "";
                    Word.Application wordApplication = new Word.Application();

                    string PathToNote = @"D:\Diplom\proga\Agent\treaty\prof\prof.doc";
                    Word.Application application = new Word.Application();
                    application.Visible = false;
                    Word.Document document = application.Documents.Open(PathToNote);
                    ReplaceWordStubs("{номер}", dataGridView1.CurrentRow.Cells[0].Value.ToString(), document);
                    ReplaceWordStubs("{дата}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);

                    string query1 = $@"Select 
position.name as [Должность]
from policyholder inner join city on policyholder.idcity=city.idcity 
inner join position on position.idposition=policyholder.idwork 
inner join work on work.idwork=position.idwork
 where idpolicyholder={dataGridView1.CurrentRow.Cells[3].Value.ToString()}";
                    DataTable data = new DataTable();
                    SqlDataAdapter command1 = new SqlDataAdapter(query1, sqlConnection);
                    command1.Fill(data);
                        DataColumn column = data.Columns[0];
                        DataRow row = data.Rows[0];
                    ReplaceWordStubs("{должность}", row[column].ToString().ToLower(), document);
                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датан}", dataGridView1.CurrentRow.Cells[9].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{датак}", dataGridView1.CurrentRow.Cells[10].Value.ToString().Substring(0, 10), document);
                    ReplaceWordStubs("{взнос}", dataGridView1.CurrentRow.Cells[17].Value.ToString(), document);
                    ReplaceWordStubs("{квзнос}", dataGridView1.CurrentRow.Cells[8].Value.ToString().ToLower(), document);

                    if (dataGridView1.CurrentRow.Cells[18].Value.ToString() != "")
                    {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10), document);
                    }
                    else if (dataGridView1.CurrentRow.Cells[20].Value.ToString() != "")
                    {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[18].Value.ToString().Substring(0, 10)+"; "+ dataGridView1.CurrentRow.Cells[20].Value.ToString().Substring(0, 10) + "; " + dataGridView1.CurrentRow.Cells[22].Value.ToString().Substring(0, 10), document);
                       
                    }
                    else
                    {
                        ReplaceWordStubs("{датав}", dataGridView1.CurrentRow.Cells[7].Value.ToString().Substring(0, 10), document);
                    }
                    ReplaceWordStubs("{страхователь}", dataGridView1.CurrentRow.Cells[4].Value.ToString(), document);
                    application.Visible = true;
                }
                    else
                    {
                    MessageBox.Show("Нет шаблона договора!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }

        }
        int visible = 0;
        private void button5_Click(object sender, EventArgs e)
        {
            panel3.Visible = false;
            panel2.Visible = false;
            panel7.Visible = false;
            panel6.Visible = false;
            //   dataGridView2.Visible = false;
            dataGridView1.Enabled = true;
            clear();
            panel2.Visible = false;
            visible = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Visible == true)
                {
                    visible++;
                }
            }
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);
            ExcelApp.Columns.NumberFormat = "General";
            ExcelWorkSheet.StandardWidth = 30;
            ExcelWorkSheet.Columns.ColumnWidth = 20;
            ExcelApp.Rows[1].Columns[6] = "Договора";
            ExcelApp.Rows[visible + 3].Columns[6] = "Ридецкая Анна Михайловна";
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                ExcelApp.Cells[2, i + 1] = dataGridView1.Columns[i].HeaderText;

            }
            for (int j = 0; j < dataGridView1.ColumnCount; j++)
            {
                for (int i = 0; i < visible; i++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Visible == true)
                    {
                        if (j == 7 || j == 9 || j == 10 || j == 16 )
                        {
                            ExcelApp.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);

                        }
                        //else if (j == 17)
                        //{
                        //    ExcelApp.Cells[i + 3, j + 1] = Convert.ToDecimal(dataGridView1.Rows[i].Cells[j].Value.ToString());

                        //}
                        else if (j == 18 && dataGridView1.Rows[i].Cells[j].Value.ToString() != "")
                        {
                            ExcelApp.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);

                        }
                        else if ((j == 20 || j == 22) && dataGridView1.Rows[i].Cells[j].Value.ToString() != "")
                        {
                            ExcelApp.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString().Substring(0, 10);

                        }
                        else
                        {
                            ExcelApp.Cells[i + 3, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();

                        }
                    }
                }

            }
            for (int i = 0; i < visible; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    Microsoft.Office.Interop.Excel.Range range = ExcelWorkSheet.Range[$"A1:X{visible + 3}"];
                    ExcelWorkSheet.Range[$"A1:X{visible + 3}"].Cells.WrapText = true;
                    range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    range.Borders.Weight = Excel.XlBorderWeight.xlMedium;
                }
            }
            ExcelApp.Columns["B"].Delete();
            ExcelApp.Columns["C"].Delete();
            ExcelApp.Columns["D"].Delete();
            ExcelApp.Columns["V"].Delete();
            //ExcelApp.Columns["E"].Delete();
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }
        int id = 0;
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            
                id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
    
        }

      
        

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        private void button10_Click(object sender, EventArgs e)
        {
           

        }

        private void button12_Click(object sender, EventArgs e)
        {
        
            if (panel7.Visible==true)
            {
                panel6.Visible = true;
                panel7.Visible = false;
                panel2.Visible = false;
                button12.Visible = false;
                
                button10.Top = 771;
                button10.Left = 1144;
               
                button11.Visible = false;
                if (id2 != 0)
                {
                    button10.Visible = true;
                }
                else
                {
                    button10.Visible = false;
                }

                for (int i = 0; i < dataGridView3.RowCount; i++)
                {
                    dataGridView3.Rows[i].Selected = false;
                }

                for (int i = 0; i < dataGridView3.RowCount; i++)
                {

                    if (id2 == Convert.ToInt32(dataGridView3[0, i].Value.ToString()))
                    {
                        dataGridView3.Rows[i].Selected = true;
                        dataGridView3.Rows[i].DefaultCellStyle.SelectionForeColor = Color.Black;
                        dataGridView3.Rows[i].DefaultCellStyle.SelectionBackColor = Color.FromArgb(212, 236, 252);
                        break;
                    }
                }

            }
            else if (panel2.Visible == true)
            {
                panel2.Visible = false;
                panel6.Visible = false;
                panel7.Visible = true;
                button12.Visible = true;
                button12.Top = 740;
                button12.Left = 403;
                button10.Visible = true;
                button10.Top = 740;
                button10.Left = 773;
                button11.Visible = false;
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
        
            if (panel6.Visible==true)
            {
               
                    panel2.Visible = false;
                    panel6.Visible=false;
                    panel7.Visible = true;
                    button12.Visible = true;
                    button12.Top = 740;
                    button12.Left = 403;
                    button10.Visible = true;
                    button10.Top = 740;
                    button10.Left = 773;
                    textBox5.Visible = true;
                    label10.Visible = true;
               

            }
            else if (panel7.Visible == true)
            {
                panel7.Visible = false;
                panel6.Visible = false;
                panel2.Visible = true;
                button12.Visible = true;
                button12.Top = 744;
                button12.Left = 76;
                button11.Visible = true;
                button11.Top = 744;
                button11.Left = 1057;
                button10.Visible = false;
            }
        }
        int count = 0;
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            count = 0;
                id2 = Convert.ToInt32(dataGridView3.CurrentRow.Cells[0].Value.ToString());
                id3= Convert.ToInt32(dataGridView3.CurrentRow.Cells[1].Value.ToString());
            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView3.CurrentRow.Cells[2].Value.ToString()== dataGridView1[6,i].Value.ToString() && Convert.ToInt32(comboBox2.SelectedValue)==Convert.ToInt32( dataGridView1[3,i].Value.ToString()))
                { count++; }
            }
            
            
            if (count==0) {button10.Visible = true;
                button10.Top = 771;
                button10.Left = 1144;
                if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Медицинское страхование")
                {
                    cof3 = Convert.ToDecimal("0,52");
                }
                else if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Страхование от несчатного случая")
                {
                    cof3 = Convert.ToDecimal("1,10");
                }
                else if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Обязательное страхование гражданской ответственности перевозчика перед пассажирами")
                {
                    cof3 = Convert.ToDecimal("1,00");
                }
                else if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Страхование от несчастных случаев пассажиров железного транспорта")
                {
                    cof3 = Convert.ToDecimal("0,50");
                }
                else if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Страхование жизни")
                {
                    cof3 = Convert.ToDecimal("1,00");
                }
                else if (dataGridView3.CurrentRow.Cells[2].Value.ToString() == "Страхование гражданской ответственности за причинение вреда в связи с осуществлением профессиональной деятельности")
                {
                    cof3 = Convert.ToDecimal("0,20");
                }
                else
                {
                    cof3 = Convert.ToDecimal("1,00");
                }
                textBox4.Text = Convert.ToString(cof3);
            }
            else
            {
                button10.Visible = false;
                MessageBox.Show("Срок действия прошлого договора еще не закончился!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        public void comboBoxvid()
        {
            try
            {
                sqlConnection.Close();
                sqlConnection.Open();
                DataSet dataSet = new DataSet();
                SqlDataAdapter command = new SqlDataAdapter($@"select bid.idbid,vid.idvida, vid.name as Название, bid.note as Пожелания from vid, policyholder,bid " +
                    $@"where bid.idvida=vid.idvida and bid.idpolicyholder=policyholder.idpolicyholder and policyholder.idpolicyholder={comboBox2.SelectedValue} and bid.status='Ожидание' ", sqlConnection);
                command.Fill(dataSet);
                dataGridView3.DataSource = dataSet.Tables[0].DefaultView;
                sqlConnection.Close();
                dataGridView3.Columns[0].Visible = false;
                dataGridView3.Columns[1].Visible = false;
                dataGridView3.AllowUserToAddRows = false;
                sqlConnection.Close();


        }
            catch { }

        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //try
            //{
                comboBoxvid();
            //}
            //catch { }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
