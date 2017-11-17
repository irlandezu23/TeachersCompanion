using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Teacherscompanion;
using System.Text.RegularExpressions;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO.Ports;
using System.Windows.Forms.DataVisualization.Charting;
using System.Management;

using Emgu;
using Emgu.CV;
using Emgu.CV.UI;
using Emgu.CV.Structure;

namespace Teacherscompanion
{
    public partial class tc : Form
    {
        public static int session_in_progress = 0;//Este 1 cand este o sesiune deschisa si 0 in rest
        public static int block_the_handler = -1;//citeste la functiile de checking
        public static int check_from_serial = 0;//asta e ca sa stie programul ca incerc sa loghez de pe seriala
        private SerialPort myport;//asta e portul meu de IO
        private String in_data;//am nevoie de un string global cand citesc de pe seriala
        public static int serial_log_flag = 0;//cu asta imi dau seama daca log-ul de pe seriala a reusit

        VideoCapture capture; //cand il initializez incepe sa preia imagini
        CascadeClassifier cc;
        Image<Bgr, byte> myimage; // frame-ul pe care il afisez
        Boolean recording = false; // true daca preia imagini
        Form ImageForm; //forma care afiseaza imagini
        ImageBox ib; // image boxul din forma care afiseaza imagini

        public tc()
        {
            InitializeComponent();
            cc = new CascadeClassifier(System.Windows.Forms.Application.StartupPath + "\\haarcascade_frontalface_default.xml");//incarc fisierul pentru face detection                                                
        }

        //Aici declar functiile necesare

        //Functia asta ia un string si transforma ultimul caracter in asci
        //returneaza conversia sirului fara ultimul caracter in int32 alipita
        //cu codul asci al ultimului caracter Ex: "441E" -> 44469
        public static Int32 CustomFormator(String s)
        {//daca formatul nu e cel bun arunca o exceptie cu mesaj util pentru utilizator
            var aux1 = s.ToCharArray();
            var aux11 = System.Convert.ToInt32(aux1[s.Length - 1]);
            var aux4=s.Remove(s.Length-1);
            var aux111 = Int32.Parse(aux4);
            if (aux11 >= 97 && aux11 <= 122) aux11 = aux11 - 32;
            if (aux111 < 100 || aux111 > 999 || aux11 < 65 || aux11 > 90)
                throw new Exception("The format is ###L!(Ex. 441E,421a)");
            else
            return aux111 * 100 + aux11;
        }

        //asta verifica daca exista seria respectiva daca da, o returneaza, daca nu arunca exceptie
        public static Int32 ClassIsReal(String s)
        {
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                var lista_clase = from aux3 in datab.Classes select aux3.Id;
                var i = CustomFormator(s);
                int aux2 = 0;
                foreach (var aux in lista_clase)
                {
                    if (aux == i) aux2 = 1;
                }
                if (aux2 == 0) throw new Exception("This class does not exists!");
                else
                {
                    return i;
                }
            }
        }
        public void FillTheClassBox(ComboBox cb)
        {//asta e functia care umple un combo box cu clasele din baza de date
            try
            {
                using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
                {
                    var a = from tclass in datab.Classes select tclass.Id;
                    foreach(var aux in a)
                    {
                        string S = (aux/100+""+(char)(aux%100));
                        cb.Items.Add(S);
                    }
                }
            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        public static void LogTheStudent(String s)
        {//functia asta logheaza studentul cu rfid-ul "s"
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    if (s.Length != 0)
                    {
                        if (session_in_progress == 0) throw new Exception("No open session!");
                        //if-ul asta e redundant dar e bun pt debugg si nu strica un nivel in
                        //plus de protectie
                        Logs ben = new Logs();
                        var i = from aux in datab.Sessions select aux.Id;
                        ben.SessionsId = i.Max();
                        //aici verific daca exista studentul pe care sa-l inregistrez
                        var lista_rfid = from aux1 in datab.Students select aux1.RFid;
                        var j = Int32.Parse(s);
                        int aux4 = 0;
                        foreach (var aux in lista_rfid)
                        {
                            if (aux == j) aux4 = 1;
                        }
                        if (aux4 == 0) throw new Exception("unknown RFID");
                        ben.StudentsRFid = j;
                        //aici verific ca sa nu fie acelasi student inregistrat de
                        //2 ori in aceeasi sesiune
                        var lista_curenta = from aux in datab.Logs
                                            where aux.SessionsId == ben.SessionsId
                                            select aux.StudentsRFid;
                        int aux3 = 0;
                        foreach (var aux in lista_curenta)
                        {
                            if (aux == j) aux3 = 1;
                        }
                        if (aux3 == 1) throw new Exception("Student already logged!");
                        //aici verific ca studentul sa fie in seria corespunzatoare sesiunii curente
                        var aux2 = from tsesiune in datab.Sessions
                                   where tsesiune.Id == ben.SessionsId
                                   select tsesiune.ClassesId;
                        var aux5 = from tstudent in datab.Students
                                   where tstudent.RFid == ben.StudentsRFid
                                   select tstudent.ClassId;
                        if (aux2.First() != aux5.First()) throw new Exception("This student is not in this class!");
                        datab.Logs.Add(ben);
                        datab.SaveChanges();
                    }
                    check_from_serial = 0;
                }
                catch (Exception exc)
                {
                    if (check_from_serial == 1)
                    {
                        check_from_serial = 0;
                        serial_log_flag = 0;
                    }
                    else
                    {
                        MessageBox.Show(exc.Message);
                        check_from_serial = 0;//redundant
                    }
                }
            }
        }

        public void CheckTheStudent(String s)
        {
            //asta checkuie studentul pe GUI, dar automat asta implica si logarea lui                 
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    if (s.Length != 0)
                    {
                        if (session_in_progress == 0) throw new Exception("No open session!");
                        for (int i = 0; i < visual_list.Items.Count; i++)
                        {
                            var aux1 = visual_list.Items[i].ToString().Split(' ');
                            String aux2 = aux1[2];
                            //MessageBox.Show(aux2);
                            if (aux2 == s)
                            {
                                int b = 0;
                                foreach (int indexChecked in visual_list.CheckedIndices)
                                {
                                    if (indexChecked == i) b = 1;
                                }
                                if (b == 0)
                                {
                                    visual_list.SetItemChecked(i, true);
                                }
                                else
                                {
                                    throw new Exception("Student already logged!");
                                }
                                return;
                            }
                        }                        
                        LogTheStudent(s);//asta o folosesc numai ca sa stiu de ce nu a mers loggul
                        serial_log_flag = 0;//daca a ajuns aici clar a crapat(sper ca e redundant)
                    }                    
                    check_from_serial = 0;
                    serial_log_flag = 0;//poate nu a luat ce trebuie de pe seriala si a crapat
                }
                catch (Exception exc)
                {
                    if (check_from_serial == 1)
                    {//daca vine de pe seriala, sa dea doar feedback la uC nu si MessageBox
                        check_from_serial = 0;
                        serial_log_flag = 0;//redundant pt "no open session" dar poade e deja logat
                    }
                    else
                    {
                        MessageBox.Show(exc.Message);
                        check_from_serial = 0;//redundant
                    }
                    
                }
            }
        }

        public static void UnlogTheStudent(String s)
        {//practic sterge log-ul unui student din ultima sesiune, daca aceasta este deschisa
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    if (s.Length != 0)
                    {
                        if (session_in_progress == 0) throw new Exception("No open session!");
                        var i = from aux in datab.Sessions select aux.Id;
                        var a = i.Max();
                        var b = from aux in datab.Logs where aux.SessionsId == a select aux.StudentsRFid;
                        var c = Int32.Parse(s);
                        int aux1 = 0;
                        foreach (var aux in b)
                        {
                            if (aux == c) aux1 = 1;
                        }
                        if (aux1 == 0) throw new Exception("The student isn't logged in the current session");
                        // Book book = (Book)bookContext.Books.Where(b => b.Id == bookId).First();
                        //bookContext.Books.Remove(book);
                        Logs tlog = (Logs)datab.Logs.Where(aux2 => (aux2.StudentsRFid == c && aux2.SessionsId == a)).First();
                        datab.Logs.Remove(tlog);
                        datab.SaveChanges();
                    }
                }
                catch(Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }
            }
        }


        //De aici incepe treaba

        private void save_student_Click(object sender, EventArgs e)
        {
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                Students ben = new Students();//nush de unde am scos numele asta pentru student...
                if (textBox1.Text.Length != 0 && textBox2.Text.Length != 0 && class_box1.Text.Length != 0 && textBox4.Text.Length != 0 )
                {
                    //Folosesc try-catch ca sa fiu sigur ca nu scriu balarii in textbox-uri
                    try
                    {
                        //Aici asignez numele si prenumele noului student
                        var k = textBox2.Text;
                        if (!Regex.IsMatch(k, @"^[\p{L}]+$")) throw new Exception("Surname must contain only letters!");
                        ben.Surname = k;
                        var l = textBox1.Text;
                        if (!Regex.IsMatch(l, @"^[\p{L}]+$")) throw new Exception("Firstname must contain only letters!");
                        ben.Firstname = l;
                        ben.ClassId = ClassIsReal(class_box1.Text);//aici verific daca exista clasa(e redundant)
                                //pentru ca dropdown box-ul oricum are ca membrii clasele din baza de date
                        //Aici am grija ca RFID-ul sa fie unic
                        var lista_rfid = from aux1 in datab.Students select aux1.RFid;
                        var j = Int32.Parse(textBox4.Text);
                        int aux4 = 0;
                        foreach (var aux in lista_rfid)
                        {
                            if (aux == j) aux4 = 1;
                        }
                        if (aux4 == 1) throw new Exception("RFID already in sistem!");
                        ben.RFid = j;
                        datab.Students.Add(ben);
                        textBox1.Clear();
                        textBox2.Clear();                        
                        class_box1.SelectedIndex = -1;
                        textBox4.Clear();
                        datab.SaveChanges();
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message);

                    }
                }
            }
        }

        private void save_class_Click(object sender, EventArgs e)
        {//asta salveaza o clasa noua
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                if (textBox5.Text.Length != 0)
                {
                    try
                    {
                        Classes ben = new Classes();
                        var lista_clase = from aux3 in datab.Classes select aux3.Id;
                        var i = CustomFormator(textBox5.Text);//asta verifica daca numele e conform
                        //cu formatul si apoi returneaza numele in format int(ca pt baza de date)
                        int aux2 = 0;
                        foreach (var aux in lista_clase)
                        {
                            if (aux == i) aux2 = 1;
                        }
                        if (aux2 == 1) throw new Exception("This class already exists!");
                        ben.Id = i;
                        datab.Classes.Add(ben);
                        textBox5.Clear();
                        datab.SaveChanges();

                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(exc.Message);
                    }
                }
            }
        }
        //asta de jos e butonul de debug(l-am folosit pentru tot felul de chestii)
        //functionalitatea sa finala in program este de a ma conecta la portul 
        //pe care vin info de la uC("Connect to module")
        private void button2_Click(object sender, EventArgs e)
        {
            //string Bport = "";
            //int Bcounter = 0;
            try
            {
                /*SelectQuery q = new SelectQuery("Win32_SerialPort");
                ManagementObjectSearcher s = new ManagementObjectSearcher(q);               
                foreach (object cur in s.Get())
                {
                    ManagementObject mo = (ManagementObject)cur;
                    object id = mo.GetPropertyValue("DeviceID");
                    object pnpId = mo.GetPropertyValue("PNPDeviceID");
                    if (pnpId.ToString().Contains("BTHENUM"))
                    {
                        Bport = id.ToString();
                        Bcounter++;
                    }
                    mo.Dispose();
                }
                s.Dispose();
                if (Bcounter == 0) throw new Exception("You don't have a Bluetooth port");
                if (Bcounter > 2) throw new Exception("You have more Bluetooth device conected!");
                */
                                
                myport = new SerialPort();
                myport.BaudRate = 9600;
                myport.PortName = textBox3.Text;
                myport.Parity = Parity.None;
                myport.DataBits = 8;
                myport.StopBits = StopBits.One;
                myport.DataReceived += Myport_DataReceived; //definesc handler-ul care
                                                            //se va actiona cand vin date de la uC pe seriala
                myport.Open();
                myport.Write("S");//dupa ce m-am conectat scriu "S" pe seriala
                                  //ca uc sa iasa din starea -1(idle)
                log_the_student.Enabled = false;
                textBox3.Text = "Conected...";
                textBox3.Enabled = false;

            }
            catch(Exception exc)
            {
                var st = new StackTrace(exc, true);//asta e pentru debugg
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show(exc.Message);
            }

            //Bucatile astea de mai jos sunt aici
            //pt ca butonul asta a avut mai multe functii
            //de-a lungul timpului
            /*check_from_serial = 1;
            CheckTheStudent(textBox4.Text);
            textBox4.Clear();*/

            //Bucata asta de mai jos o tin doar de siguranta, am facut o functie care
            //ia un string si face prezenta pentru studentul cu rfid-ul respectiv
            /*using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    if(textBox4.Text.Length!=0)
                    {
                        if (session_in_progress == 0) throw new Exception("No open session!");
                        //if-ul asta e redundant dar e bun pt debugg si nu strica un nivel in
                        //plus de protectie
                        Logs ben = new Logs();
                        var i = from aux in datab.Sessions select aux.Id;
                        ben.SessionsId = i.Max();
                        //aici verific daca exista studentul pe care sa-l inregistrez
                        var lista_rfid = from aux1 in datab.Students select aux1.RFid;
                        var j = Int32.Parse(textBox4.Text);
                        int aux4 = 0;
                        foreach (var aux in lista_rfid)
                        {
                            if (aux == j) aux4 = 1;
                        }
                        if (aux4 == 0) throw new Exception("unknown RFID");
                        ben.StudentsRFid = j;
                        //aici verific ca sa nu fie acelasi student inregistrat de
                        //2 ori in aceeasi sesiune
                        var lista_curenta = from aux in datab.Logs
                                            where aux.SessionsId == ben.SessionsId
                                            select aux.StudentsRFid;
                        int aux3 = 0;
                        foreach(var aux in lista_curenta)
                        {
                            if (aux == j) aux3 = 1;
                        }
                        if (aux3 == 1) throw new Exception("Student already logged!");
                        //aici verific ca studentul sa fie in seria corespunzatoare sesiunii curente
                        var aux2 = from tsesiune in datab.Sessions
                                   where tsesiune.Id == ben.SessionsId
                                   select tsesiune.ClassesId;
                        var aux5 = from tstudent in datab.Students
                                   where tstudent.RFid == ben.StudentsRFid
                                   select tstudent.ClassId;
                        if (aux2.First() != aux5.First()) throw new Exception("This student is not in this class!");
                        datab.Logs.Add(ben);
                        datab.SaveChanges();
                        textBox4.Clear();
                        

                    }
                }
                catch(Exception exc)
                {
                    MessageBox.Show(exc.Message);
                    textBox4.Clear();
                }
            }*/
        }

        private void Myport_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                in_data = myport.ReadLine();
                int in_data_int = Int32.Parse(in_data);
                if (session_in_progress == 1)
                {
                    check_from_serial = 1;
                    serial_log_flag = 1;                    
                    this.Invoke(new EventHandler(display_uid));//trimite actiunea unui handler 
                                                               //de pe thread-ul principal
                    if (serial_log_flag == 1)
                    {//daca a mers log-ul, trimite "K" pe seriala
                        myport.Write("K");
                    }
                    else
                    {//daca nu, trimite "N" pe seriala
                        myport.Write("N");
                    }
                    serial_log_flag = 0;                    
                }
                else
                {
                    myport.Write("K");//daca nu-i sesiunea deschisa, trimite automat "K"
                    this.Invoke(new EventHandler(display_uid));                    
                }

            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void display_uid(object sender, EventArgs e)
        {
            if (session_in_progress == 1)
            {//daca e sesiunea deschisa incerc sa loghez studentul
                int in_data_int = Int32.Parse(in_data);//aici in parsez de doua ori (string->int) si
                CheckTheStudent(in_data_int.ToString());//(int->string) pentru ca pe seriala vine cu
                            //un " " in coada care ma deranjeaza(prin parsarile astea scap de spatiu)
            }
            else
            {//daca nu e sesiunea deschisa, scriu uid-ul in textBox4 pentru o eventuala inscriere
                textBox4.Text = in_data;
            }
        }

        private void session_button_Click(object sender, EventArgs e)
        {
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    switch (session_in_progress)
                    {
                        case 0://deschid o sesiune
                            if (class_box2.Text.Length != 0)
                            {
                                Sessions ben = new Sessions();
                                ben.Date = DateTime.Now;
                                ben.ClassesId = ClassIsReal(class_box2.Text);
                                datab.Sessions.Add(ben);
                                datab.SaveChanges();
                                textBox5.Clear();
                                class_box2.SelectedIndex = -1;
                                session_button.Text = "End Session";
                                session_in_progress = 1;
                                checker_button.Enabled = true;
                                save_class.Enabled = false;
                                save_student.Enabled = false;
                                class_box2.Enabled = false;
                                textBox5.Enabled = false;
                                delete_class.Enabled = false;
                                textBox2.Enabled = false;
                                class_box1.Enabled = false;
                                textBox1.Enabled = false;                                
                                //textBox3.Enabled = false;
                                tabPage2.Enabled = false;
                                tabPage3.Enabled = false;
                                //textBox4.Enabled = false;
                                show_statistic.Enabled = false;
                                //log_the_student.Enabled = false;
                                var students_in_class = from tstudent in datab.Students
                                                        where tstudent.ClassId == ben.ClassesId
                                                        orderby tstudent.Surname
                                                        select tstudent;
                                foreach(var tstudent in students_in_class)
                                {
                                    visual_list.Items.Add(tstudent.Surname +" "+ tstudent.Firstname +" "+
                                        tstudent.RFid);
                                }
                                                           
                            }
                            else
                            {
                                throw new Exception("You need a Class Name to start a session!");
                            }


                            break;
                        case 1://inchid sesiunea(ies din starea interna a sistemului "sesiune deschisa")
                            session_in_progress = 0;
                            session_button.Text = "Start Session";
                            textBox4.Clear();//redundant, doar de frumusete
                            //textBox4.Enabled = true;
                            //textBox3.Enabled = true;
                            checker_button.Enabled = false;
                            save_class.Enabled = true;
                            save_student.Enabled = true;
                            tabPage2.Enabled = true;
                            tabPage3.Enabled = true;
                            textBox5.Clear();
                            class_box2.SelectedIndex = -1;
                            class_box2.Enabled = true;
                            textBox5.Enabled = true;
                            textBox2.Enabled = true;
                            delete_class.Enabled = true;
                            class_box1.Enabled = true;
                            textBox1.Enabled = true;
                            show_statistic.Enabled = true;
                            //log_the_student.Enabled = true;
                            visual_list.Items.Clear();
                            break;
                    }
                }
                catch(Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }

            }
        }

        private void show_statistic_Click(object sender, EventArgs e)
        {//practic ce fac este sa creez un tabel mare de forma care ma intereseaza,
            //iar apoi sa-l export intr-un excel
            using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
            {
                try
                {
                    if (class_box3.Text.Length != 0)
                    {
                    //aici declar tabelul(creez coloanele de care am nevoie)
                    System.Data.DataTable statistic = new System.Data.DataTable(class_box3.Text);
                    DataColumn nume = statistic.Columns.Add("Surname", typeof(String));
                    DataColumn prenume = statistic.Columns.Add("Firstname", typeof(String));
                    DataColumn rfid = statistic.Columns.Add("RFID", typeof(Int32));
                    var aux1 = ClassIsReal(class_box3.Text);
                    var aux2 = from tsession in datab.Sessions
                               where tsession.ClassesId == aux1
                               select tsession.Date;
                    foreach (var aux in aux2)
                    {                        
                        statistic.Columns.Add(aux.ToString(), typeof(String));
                    }
                    //aici scot numele, prenumele si RFID-ul studentilor din seria respectiva
                    //sunt ordonate toate dupa surname ca sa nu difere indicii
                    var aux3 = from tstudent in datab.Students
                               where tstudent.ClassId == aux1 // aux1 contine numele clasei
                               orderby tstudent.Surname
                               select tstudent.Surname;
                    var aux4 = from tstudent in datab.Students
                               where tstudent.ClassId == aux1
                               orderby tstudent.Surname
                               select tstudent.Firstname;
                    var aux5 = from tstudent in datab.Students
                               where tstudent.ClassId == aux1
                               orderby tstudent.Surname
                               select tstudent.RFid;

                        //aici declar fiecare linie in parte care va avea NUME|PRENUME|RFID|P/A|P/A|...
                        //P/A in functie de Prezent/Absent
                        for (int i = 0; i < aux5.Count(); i++)
                        {
                            //aici scot logurile unui student
                            var aux8 = aux5.ToArray()[i];
                            var aux6 = from tlog in datab.Logs
                                       where tlog.StudentsRFid == aux8
                                       select tlog;
                            var aux7 = aux6.ToArray();
                            //aici scot datetime-ul din sesiunea corespunzatoare logului
                            DateTime[] a = new DateTime[aux7.Length+10];//asta asigura o alocare dinamica
                            int k = 0;
                            foreach (var aux in aux7)
                            {//umplu variabila a cu toate datele logurilor unui student
                                a[k++] = aux.Session.Date;
                            }
                            //aici creez efectiv randul si incep sa-l umplu
                            DataRow trow = statistic.NewRow();
                            var columns = statistic.Columns.Cast<DataColumn>();
                            k = 0;
                            foreach (var aux in columns)
                            {
                                if (aux.ColumnName == "Surname")
                                    trow[k++] = aux3.ToArray()[i];
                                else
                                {
                                    if (aux.ColumnName == "Firstname")
                                        trow[k++] = aux4.ToArray()[i];
                                    else
                                    {
                                        if (aux.ColumnName == "RFID")
                                            trow[k++] = aux5.ToArray()[i];
                                        else
                                        {
                                            int h = 0;
                                            foreach (var tdate in a)//a e variabila cu date
                                            {//daca in datele sesiunilor in care a fost logat studentul se regaseste
                                                //data coloanei atunci studentul e prezent
                                                if (aux.ColumnName == tdate.ToString()) h = 1;
                                            }
                                            if (h == 1) trow[k++] = "Prezent";
                                            else
                                            {
                                                trow[k++] = "Absent";
                                            }

                                        }
                                    }
                                }



                            }
                            statistic.Rows.Add(trow);//dupa ce am completat randul il pun in tabel
                        }                        
                        class_box3.SelectedIndex = -1;
                        //export to excel
                        if (true)
                        {
                            Microsoft.Office.Interop.Excel.Application oXL;
                            Microsoft.Office.Interop.Excel._Worksheet oSheet;
                            Microsoft.Office.Interop.Excel._Workbook oWB;
                            Microsoft.Office.Interop.Excel.Range oRange;
                            oXL = new Microsoft.Office.Interop.Excel.Application();
                            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
                            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                            oXL.Visible = true;
                            var columns = statistic.Columns.Cast<DataColumn>();
                            var rows = statistic.Rows.Cast<DataRow>();
                            //asta e pentru casuta unde se afla numele seriei
                            oSheet.Cells[1, 1] = statistic.TableName;
                            oSheet.Cells[1, 1].Font.Bold = true;
                            oSheet.Cells[1, 1].Borders.Weight = 3.2;
                            oSheet.Cells[1, 1].Borders.Color = Color.Black;
                            oSheet.Cells[1, 1].Borders.LineStyle = XlLineStyle.xlContinuous;

                            object clr;

                            //foreach-ul asta e pentru header-ul tabelei
                            int k = 2;
                            foreach (var auxcolumn in columns)
                            {                                
                                oSheet.Cells[2, k] = (auxcolumn.ColumnName+"");
                                oRange = oSheet.Cells[2, k];
                                oSheet.Cells[2, k].Font.Bold = true;
                                oSheet.Cells[2, k].Borders.Weight = 3.2;
                                oSheet.Cells[2, k].Borders.Color = Color.Black;
                                oSheet.Cells[2, k].Borders.LineStyle = XlLineStyle.xlContinuous;
                                clr = oRange.Interior.Color;
                                clr = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                                oRange.Interior.Color = clr;
                                oSheet.Columns[k].ColumnWidth = 18;//ca sa fiu sigur ca incape tot ce trebuie
                                //asta de jos e pentru randul pe care scrie nr sesiunii
                                if (k > 4)
                                {
                                    oSheet.Cells[1, k] = ("Sesiunea" +" "+ (k - 4));
                                    oSheet.Cells[1, k].Font.Bold = true;
                                    oSheet.Cells[1, k].Borders.Weight = 3.2;
                                    oSheet.Cells[1, k].Borders.Color = Color.Black;
                                    oSheet.Cells[1, k].Borders.LineStyle = XlLineStyle.xlContinuous;
                                    oRange = oSheet.Cells[1, k];
                                    oRange.Interior.Color = clr;
                                }
                                k++;
                            }
                            if (k > 5)//asta creeaza header-ul pentru coloana de procentaje
                            {
                                oSheet.Cells[1, k] = ("Prezenta");
                                oSheet.Cells[1, k].Font.Bold = true;
                                oSheet.Cells[1, k].Borders.Weight = 3.2;
                                oSheet.Cells[1, k].Borders.Color = Color.Black;
                                oSheet.Cells[1, k].Borders.LineStyle = XlLineStyle.xlContinuous;
                                oRange = oSheet.Cells[1, k];
                                clr = oRange.Interior.Color;
                                clr = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                                oRange.Interior.Color = clr;
                            }
                            //asta e pentru restul tabelei
                            int counter;
                            int h = 3;
                            foreach (var auxraw in rows)
                            {
                                k = 2;
                                counter = 0;
                                foreach (var auxcolumn in columns)
                                {
                                    var telem = statistic.Rows[h-3].ItemArray.Select(x => x.ToString()).ToArray();
                                    oSheet.Cells[h, k] = telem[k - 2];
                                    oSheet.Cells[h, k].Borders.Weight = 3.2;
                                    oSheet.Cells[h, k].Borders.Color = Color.Black;
                                    oSheet.Cells[h, k].Borders.LineStyle = XlLineStyle.xlContinuous;
                                    
                                    if (telem[k - 2] == "Absent")
                                    {
                                        oRange = oSheet.Cells[h, k];
                                        clr = oRange.Interior.Color;
                                        clr = System.Drawing.ColorTranslator.ToOle(Color.OrangeRed);
                                        oRange.Interior.Color = clr;
                                    }
                                    else
                                    {
                                        if (telem[k - 2] == "Prezent")
                                        {
                                            oRange = oSheet.Cells[h, k];
                                            clr = oRange.Interior.Color;
                                            clr = System.Drawing.ColorTranslator.ToOle(Color.LightGreen);
                                            oRange.Interior.Color = clr;
                                            counter++;//pentru a face procentajele la final
                                        }
                                        else
                                        {
                                            oRange = oSheet.Cells[h, k];
                                            clr = oRange.Interior.Color;
                                            clr = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                                            oRange.Interior.Color = clr;
                                        }
                                    }
                                    k++;

                                }
                                if (k > 5) // dupa ce am completat tabela, va exista la final o coloana
                                {//unde scrie procentajul (k>5=>cel putin o sesiune)
                                    oSheet.Cells[h, k] = ((counter * 100 /(k-5))+"%");
                                    oSheet.Cells[h, k].Font.Bold = true;
                                    oSheet.Cells[h, k].Borders.Weight = 3.2;
                                    oSheet.Cells[h, k].Borders.Color = Color.Black;
                                    oSheet.Cells[h, k].Borders.LineStyle = XlLineStyle.xlContinuous;
                                    oRange = oSheet.Cells[h, k];
                                    clr = oRange.Interior.Color;
                                    clr = System.Drawing.ColorTranslator.ToOle(Color.LightGray);
                                    oRange.Interior.Color = clr;
                                }
                                h++;
                            }
                        }
                        
                        

                    }
                    else
                    {
                        throw new Exception("Select a class first!");
                    }

                }
                catch(Exception exc)
                {
                    var st = new StackTrace(exc, true);//asta e pentru debugg
                    var frame = st.GetFrame(0);
                    var line = frame.GetFileLineNumber();
                    MessageBox.Show(exc.Message+"  "+line);                    
                    class_box3.SelectedIndex = -1;
                }
            }
        }

        private void visual_list_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            try
            {
                if (block_the_handler!=-1)//daca am dat no la checkbox handler-ul s-ar apela dinou
                {//if-ul asta blocheaza actionarea handlerului de doua ori
                    block_the_handler = -1;
                    return;
                }
                else
                {
                    var a = visual_list.Items[e.Index].ToString();
                    var b = e.CurrentValue.ToString();
                    var c = a.Split(' ');
                    DialogResult dialogResult;
                    if (check_from_serial==0) //practic daca vine de pe seriala log-ul evit dialog boxul
                    {
                        dialogResult = MessageBox.Show(
                            "You are about to change the status of the student: " + c[0] + " " + c[1]
                            + "\n Do you want to continue? ", "Caution",
                            MessageBoxButtons.YesNo);
                    }
                    else
                    {
                        dialogResult = DialogResult.Yes;//e ca si cum dialog box-ul ar returna Yes                        
                    }
                    
                    
                                        
                    if (dialogResult == DialogResult.Yes)
                    {
                        switch (b)
                        {
                            case "Checked":
                                {
                                    if (check_from_serial == 0)
                                    {
                                        UnlogTheStudent(c[2]);
                                    }
                                    else
                                    {
                                        check_from_serial=0;
                                        serial_log_flag = 0;//e logat, nu ma intereseaza sa-l deloghez cu cardul
                                    }
                                    break;
                                }
                            case "Unchecked":
                                {
                                    LogTheStudent(c[2]);
                                    break;
                                }
                            default: throw new Exception("Something went wrong!");
                        }
                    }
                    else
                    {
                        block_the_handler = e.Index;
                        /* de aici trimit indexul intr-o variabila globala in cazul
                         * in care s-a apasat NO de catre utilizator. Functia care se executa dupa triggerarea
                         * eventului va lua index-ul, v-a vedea daca casuta s-a checkuit/decheckuit, va face
                         * actiunea inversa pentru ca s-a apasat NO. Apoi acest handler se va actiona dinou
                         * deoarece s-a schimbat dinou starea, dar if-ul de la inceput are grija sa ii taie 
                         * functionalitatea si sa reinitializeze valoarea lui "block_the_handler" la valoarea
                         * default, adica -1 pentru a nu strica functionarea handlerului pe viitor
                         */
                    }
                }
                
            }
            catch(Exception exc)
            {
                var st = new StackTrace(exc, true);//asta e pentru debugg
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show(exc.Message + "  " + line);                
            }
        }

        private void visual_list_SelectedIndexChanged(object sender, EventArgs e)
        {//asta actioneaza numai daca ai dat "no" cand loguiesti din GUI
            try
            {//practic checked value se schimba de doua ori dar a doua oara se blocheaza handler-ul de mai sus
                if (block_the_handler != -1)
                {
                    int b = 0;
                    foreach (int indexChecked in visual_list.CheckedIndices)
                    {
                        if (indexChecked == block_the_handler) b = 1;
                    }
                    switch (b)
                    {
                        case 1:
                            {
                                visual_list.SetItemChecked(block_the_handler, false);
                                break;
                            }
                        case 0:
                            {
                                visual_list.SetItemChecked(block_the_handler, true);
                                break;
                            }
                        default: throw new Exception("Something went wrong!");
                    }
                }
            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void class_box1_DropDown(object sender, EventArgs e)
        {//de fiecare data cand deschid dropdown-ul, asta se umple cu clasele existente in baza de date
            class_box1.Items.Clear();
            FillTheClassBox(class_box1);
        }

        private void class_box2_DropDown(object sender, EventArgs e)
        {
            class_box2.Items.Clear();
            FillTheClassBox(class_box2);
        }

        private void delete_class_Click(object sender, EventArgs e)
        { //deleteaza log-urile studentilor din clasa, apoi studentii, apoi sesiunile clasei, apoi clasa insusi
            try
            {
                using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
                {
                    if (class_box2.Text.Length == 0) throw new Exception("Select a class first!");                  
                    var a = class_box2.Text;
                    var b = CustomFormator(a);
                    DialogResult dialogResult = MessageBox.Show(
                       "You are about to permanently delete class: " + a +"\n"+
                       "That means all logs, sessions and students associated with this class will be removed"+
                       "from the internal database. We kindly recommend you to save the statistics first!"
                       + "\n Are you sure you want to delete this class? ", "Caution",
                       MessageBoxButtons.YesNo);

                    if (dialogResult == DialogResult.Yes)
                    {
                        var aux1 = from aux in datab.Sessions where aux.ClassesId == b select aux;
                        foreach (var tsession in aux1)
                        {
                            var aux2 = from aux in datab.Logs where aux.SessionsId == tsession.Id select aux;
                            foreach (var tlog in aux2)
                            {
                                datab.Logs.Remove(tlog);
                            }
                            datab.Sessions.Remove(tsession);
                        }
                        var aux3 = from aux in datab.Students where aux.ClassId == b select aux;
                        foreach (var tstudent in aux3)
                        {
                            datab.Students.Remove(tstudent);
                        }
                        var aux4 = from aux in datab.Classes where aux.Id == b select aux;
                        var tclass = aux4.First();
                        datab.Classes.Remove(tclass);
                        datab.SaveChanges();
                        class_box2.SelectedIndex = -1; //chestia asta face ca dropdown boxul sa nu afiseze nimic
                    }
                    else
                    {
                        class_box2.SelectedIndex = -1;
                        return;
                    }
                }
            }
            catch(Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void tc_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                myport.Write("E");//cand inchizi aplicatia trimite E ca uC sa intre in starea -1(IDLE)
            }
            catch(Exception)
            {
                ;//in caz ca nu e conectat nici un modul o sa arunce exceptie, dar nu facem nimic cu ea
            }
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (!e.TabPage.Enabled)
            {
                e.Cancel = true;
            }
        }

        private void class_box3_DropDown(object sender, EventArgs e)
        {
            class_box3.Items.Clear();
            FillTheClassBox(class_box3);
        }

        private void chart_button_Click(object sender, EventArgs e)
        {
            try
            {
                using (Teacherscompanion.TCdbmodelContainer datab = new TCdbmodelContainer())
                {                                                   
                    chart1.Series["Activity"].Points.Clear();
                    string class_string = class_box3.Text;
                    //class_box3.SelectedIndex = -1;
                    int class_int = CustomFormator(class_string);
                    var ben = from aux in datab.Students
                               where aux.ClassId == class_int
                               orderby aux.Surname descending
                               select aux;                    
                    var sesiuni = from aux in datab.Sessions
                                where aux.ClassesId == class_int
                                select aux.Id;
                    var sesiuni_count = sesiuni.Count();
                    if (sesiuni_count == 0) sesiuni_count = 1;//ca sa evit impartirea la 0
                            //oricum daca sunt 0 sesiuni studentii au 0 log-uri
                    foreach(var aux in ben)
                    {
                        var loguri = from tlog in datab.Logs
                                     where tlog.StudentsRFid == aux.RFid
                                     select tlog.Id;
                        var log_count = loguri.Count();
                        var procentaj = log_count * 100 / sesiuni_count;
                        string afisaj = (aux.Surname+" "+aux.Firstname);
                                                
                        chart1.Series["Activity"].Points.AddXY(afisaj, procentaj);
                        
                    }

                    var chartArea = chart1.ChartAreas["ChartArea1"];
                    chartArea.AxisX.MajorGrid.LineWidth = 0;
                    chartArea.AxisY.MajorGrid.LineWidth = 0;
                    chartArea.AxisY.Minimum = 0;
                    chartArea.AxisY.Maximum = 100;
                    chartArea.AxisX.Minimum = 0;
                    chartArea.AxisX.Maximum = ben.Count()+1;
                    chartArea.AxisX.Interval = 1;
                    chartArea.AxisY.Interval = 10;                   
                    //chartArea.CursorX.AutoScroll = true;
                    //chartArea.AxisX.ScaleView.Zoomable = true;
                    //chartArea.AxisX.ScaleView.SizeType = DateTimeIntervalType.Number;
                    int position = 0; //de unde incepe afisarea
                    int size = 7; //cate incap
                    chartArea.AxisX.ScaleView.Zoom(position,size);
                    chartArea.AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll;
                    //chartArea.AxisX.
                    //chartArea.AxisX.ScaleView.SmallScrollSize = 10;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void textBox3_DropDown(object sender, EventArgs e)
        {
            textBox3.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                textBox3.Items.Add(port);
            }
        }

        private void checker_button_Click(object sender, EventArgs e)
        {
            try
            {
                if (!recording)
                {
                    checker_button.Text = "Stop Camera";
                    session_button.Enabled = false;
                    ImageForm = new Form(); //creez forma
                    ImageForm.ControlBox = false; //scot butoanele ca nu cumva sa-mi iau vreun crash
                    ImageForm.Visible = true;
                    ImageForm.Name = "Class View"; // e cam degeaba
                    System.Drawing.Point location = visual_list.PointToScreen(System.Drawing.Point.Empty);//locatia la care ma raportez pentru forma a doua
                    location.X = location.X + 200;
                    location.Y = location.Y - 125;
                    ImageForm.Location = location;
                    ImageForm.AutoSizeMode = AutoSizeMode.GrowAndShrink;//forma sa fie cat image box
                    ImageForm.AutoSize = true;
                    ImageForm.MaximizeBox = false; // si asta e cam degeaba
                    ib = new ImageBox();
                    ib.FunctionalMode = ImageBox.FunctionalModeOption.Minimum;//ca sa nu mearga zoom si prostii
                    ib.Location = new System.Drawing.Point(1, 1);
                    ib.Size = new System.Drawing.Size(600, 350);
                    ib.SizeMode = PictureBoxSizeMode.StretchImage;//adapteaza poza la image box
                    ImageForm.Controls.Add(ib);
                    capture = new VideoCapture();//dau drumul la inregistrare
                    recording = true;
                    backgroundWorker1.RunWorkerAsync();//apelez background worker-ul
                }
                else
                {
                    checker_button.Text = "Start Camera";
                    session_button.Enabled = true;
                    recording = false;
                    ImageForm.Dispose();//sa nu ocup memoria sau sa apara probleme
                }
            }
            catch (Exception exc)
            {
                var st = new StackTrace(exc, true);//asta e pentru debugg
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show(exc.Message + "  " + line);
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                int maxnr = 0;//max de fete detectate
                while (recording)
                {
                    myimage = capture.QueryFrame().ToImage<Bgr, byte>();//ia imaginea
                    if (myimage != null)
                    {
                        var grayframe = myimage.Convert<Gray, byte>();//o face grayframe pentru algoritmul de detectie
                        var faces = cc.DetectMultiScale(grayframe, 1.1, 10, Size.Empty); //aici face detectia
                        foreach (var face in faces)
                        {
                            myimage.Draw(face, new Bgr(Color.LightCoral), 3); //pune dreptunghiuri pe fete pe imaginea afisata

                        }
                        if (maxnr < faces.Length)
                        {
                            maxnr = faces.Length; // daca am detectat mai multe schimb maximul
                        }
                    }
                    //imageBox1.Image = myimage;
                    //imageBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    ib.Image = myimage; // afisez imaginea preluata si cu dreptunghiuri in image box
                }
                capture.Dispose(); //eliberez memoria sunt sigur ca nu iau crash
                var checked_students = visual_list.CheckedIndices.Count;//numara cati am checkuiti
                MessageBox.Show("You have "+checked_students+" checked students and "+maxnr+" video detected students!");
            }
            catch (Exception exc)
            {
                var st = new StackTrace(exc, true);//asta e pentru debugg
                var frame = st.GetFrame(0);
                var line = frame.GetFileLineNumber();
                MessageBox.Show(exc.Message + "  " + line);
            }
        }
    }
}
