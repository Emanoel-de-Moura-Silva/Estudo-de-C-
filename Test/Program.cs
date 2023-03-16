using System;
using System.Diagnostics;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //Worbook e planilha por inteiro//
            using(var workbook = new XLWorkbook()){
                var worksheet = workbook.Worksheets.Add("Planilha1");
                worksheet.Cell("A1").Value = "Ola mundo";
                worksheet.Cell("A2").Value = 10;
                worksheet.Cell("A3").Value = 26;
                worksheet.Cell("A4").Value = 14;

                //Fazendo formula no Excell com c#//
                //FormulaA1 Seleciona ate onde voce quer, FormulaR1C1 Seleciona linha e coluna//
                //Colocar a formula do Excell em Inglês//
                worksheet.Cell("A5").FormulaA1 = "SUM(A2:A4)";

                //Ultilizando Imagem
                var imagem = @"C:\Users\B901175\temp\Imagem.JFIF";
                worksheet.AddPicture(imagem).MoveTo(worksheet.Cell("A10")).Scale(0.5);

                //Ultilizando bordas
                worksheet.Cell("A1").Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Cell("A1").Style.Border.BottomBorderColor = XLColor.Blue;

                //Calculo com a formula mostrando quando executa sem precisar iniciar o excel
                System.Console.WriteLine("Valor da Soma: {0}", worksheet.Cell("A5").Value);



                workbook.SaveAs(@"C:\Users\B901175\temp\testeExcel.xlsx");
                

            

            }
            //Process.Start(new ProcessStartInfo(@"C:\Users\B901175\temp\testeExcel.xlsx") { UseShellExecute = true});
        }
    }
}
