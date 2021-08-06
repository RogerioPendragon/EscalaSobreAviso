using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using Microsoft.Office.Interop.Excel;

namespace Escalasobreaviso
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = excel.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];
            Console.Write("Digite a quantidade de plantonistas: ");
            int plantonista = Convert.ToInt32(Console.ReadLine());
            string[] matriculas;
            string[] nome;
            string plantao = "00:00 - 08:59 / 19:01 - 23:59";
            string escala = "09:00 - 19:00";
            nome = new string[plantonista];
            matriculas = new string[plantonista];
            int indice = 1;
            for (int i=0; i<=plantonista-1;i++)
            {
                Console.Write("Digite o nome do " + indice + "º plantonista: ");
                nome[i] = Console.ReadLine();
                nome[i] =nome[i].ToUpper(new CultureInfo("tr-TR", false));
                Console.Write("Digite a matricula do " + indice + "º plantonista: ");
                matriculas[i] = "019900" + Console.ReadLine();
                indice = indice + 1;
            }
            Console.Write("Digite a data Inicial: ");
            DateTime dataini = Convert.ToDateTime(Console.ReadLine());
            Console.Write("Digite a data Final: ");
            DateTime datafim = Convert.ToDateTime(Console.ReadLine());
            ws.Cells[1, 1] = "DATA";
            ws.Cells[1, 2] = "DIA DA SEMANA";
            ws.Cells[1, 3] = "NOME";
            ws.Cells[1, 4] = "MATRICULA";
            ws.Cells[1, 5] = "HORÁRIO PLANTÃO";
            ws.Cells[1, 6] = "HORÁRIO ESCALA RÍGIDA";
            int contador = 0;
            int indiceplanilha = 2;
            while (dataini <= datafim)
            {
                int diasemana = (int)dataini.DayOfWeek;
                string nomediasemana = new CultureInfo("pt-BR").DateTimeFormat.GetDayName((DayOfWeek)diasemana);
                nomediasemana = nomediasemana.ToUpper(new CultureInfo("tr-TR", false));
                string data = dataini.ToString("MM/dd/yyyy");
                if (diasemana ==0 ) 
                {
                    ws.Cells[indiceplanilha, 1]= data;
                    ws.Cells[indiceplanilha, 2] = nomediasemana;
                    ws.Cells[indiceplanilha, 3] = nome[contador];
                    ws.Cells[indiceplanilha, 4] = matriculas[contador];
                    ws.Cells[indiceplanilha, 5] = "Integral";
                    ws.Cells[indiceplanilha, 6] = escala;
                    indiceplanilha = indiceplanilha + 1;
                    contador = contador + 1;
                    if (contador> plantonista-1)
                    {
                        contador = 0;
                    }
                }
                else if (diasemana == 6)
                {
                    //ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 2] = nomediasemana;
                    ws.Cells[indiceplanilha, 3] = nome[contador];
                    ws.Cells[indiceplanilha, 4] = matriculas[contador] ;
                    ws.Cells[indiceplanilha, 5] = "Integral";
                    ws.Cells[indiceplanilha, 6] = escala;
                    indiceplanilha = indiceplanilha + 1;
                }
                else if (diasemana==1)
                {
                    //ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 2] = nomediasemana;
                    ws.Cells[indiceplanilha, 3] = nome[contador];
                    ws.Cells[indiceplanilha, 4] = matriculas[contador] ;
                    ws.Cells[indiceplanilha, 5] = plantao;
                    ws.Cells[indiceplanilha, 6] = escala;
                    indiceplanilha = indiceplanilha + 1;
                }
                else
                {
                    //ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 1] = data;
                    ws.Cells[indiceplanilha, 2] = nomediasemana;
                    ws.Cells[indiceplanilha, 3] = nome[contador];
                    ws.Cells[indiceplanilha, 4] = matriculas[contador] ;
                    ws.Cells[indiceplanilha, 5] = plantao;
                    ws.Cells[indiceplanilha, 6] = escala;
                    indiceplanilha = indiceplanilha + 1;
                    //Console.WriteLine(data + "---" + nomediasemana + "---" + nome[contador] + "---" + matriculas[contador] + "---" + plantao + "---" + escala);
                }
                dataini = dataini.AddDays(1);
            }
            excel.Visible = true;
            //Console.ReadLine();
        }
    }
}
