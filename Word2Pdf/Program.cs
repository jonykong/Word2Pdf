using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
/*using System.IO;*/
using System.Windows.Forms;

namespace Word2Pdf
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
//         static void Main(string[] args)
//         {
//             if (args.Length != 1)
//             {
//                 PrintUsage();
//                 return;
//             }
// 
//             var path = args[0];
//             if (Directory.Exists(path))
//             {
//                 Run(path);
//             }
//             else if(path.EndsWith(".docx") || path.EndsWith(".doc"))
//             {
//                 Process(path);
//             }
//             else
//             {
//                 Console.WriteLine("{0} is not a directory");
//             }
//         }

//         private static void PrintUsage()
//         {
//             Console.WriteLine("Pass the directory as a param");
//         }

//         public static void Run(string path)
//         {
//             var files = Directory.GetFiles(path, "*.docx").Union(Directory.GetFiles(path, "*.doc"));
//             foreach (var file in files)
//             {
//                 if (!Path.GetFileName(file).StartsWith("~"))
//                 {
//                     Process(file);
//                 }
//             }
// 
//             var subs = Directory.GetDirectories(path);
//             foreach (var sub in subs)
//             {
//                 Run(sub);
//             }
//         }

//         public static void Process(string file)
//         {
//             Console.WriteLine("Found: {0}", file);
//             File.SetAttributes(file, FileAttributes.Normal);
//             var idx = file.LastIndexOf(".");
//             var newPath = file.Remove(idx, file.Length - idx) + ".pdf";
//             Doc2PDFAtServerClass.word2PdfFcih(file, newPath);
//         }
    }
}
