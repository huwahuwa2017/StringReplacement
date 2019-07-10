using ClosedXML.Excel;
using Mono.Cecil;
using Mono.Cecil.Cil;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace StringReplacement
{
    public class Program
    {
        private static int SuccessCount;

        private static int ErrorCount;



        [STAThread]
        private static void Main(string[] args)
        {
            string DirectoryName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string DLLPath = @"C:\Program Files (x86)\Steam\steamapps\common\From The Depths\From_The_Depths_Data\Managed\Ftd";
            string InputPath = DirectoryName;



            string Text0 = "DLLファイルを選択してください";
            Console.WriteLine(Text0);

            OpenFileDialog OFD = new OpenFileDialog
            {
                Title = Text0,
                InitialDirectory = Path.GetDirectoryName(DLLPath),
                FileName = DLLPath,
                Filter = "dll files (*.dll)|*.dll|All files (*.*)|*.*"
            };

            if (OFD.ShowDialog() == DialogResult.OK)
            {
                DLLPath = OFD.FileName;
            }

            Console.WriteLine(DLLPath);
            OFD.Dispose();



            string Text1 = "翻訳ファイルを指定してください";
            Console.WriteLine(Text1);

            OpenFileDialog SFD = new OpenFileDialog
            {
                Title = Text1,
                InitialDirectory = Path.GetDirectoryName(InputPath),
                FileName = InputPath,
                Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
            };

            if (SFD.ShowDialog() == DialogResult.OK)
            {
                InputPath = SFD.FileName;
            }

            Console.WriteLine(InputPath);
            SFD.Dispose();



            Console.WriteLine("キー入力で文字列の置き換えを開始します");
            Console.ReadKey();



            foreach (string FilePath in Directory.GetFiles(Path.GetDirectoryName(DLLPath), "*.dll"))
            {
                string FileName = Path.GetFileName(FilePath);
                Console.WriteLine(FileName + "の取得中");

                try
                {
                    File.Copy(FilePath, Path.Combine(DirectoryName, FileName), true);
                }
                catch
                {
                    Console.WriteLine("使用中のファイル : " + FileName);
                }
            }

            AssemblyDefinition AssemblyDef = AssemblyDefinition.ReadAssembly(DLLPath, new ReaderParameters { ReadWrite = true });
            IEnumerable<TypeDefinition> TypeDefList = AssemblyDef.Modules.SelectMany(x => x.Types);

            IXLWorksheet IXLW = new XLWorkbook(InputPath).Worksheet("Translation");
            int RowNum = IXLW.LastRowUsed().RowNumber();
            int LineCount = 0;

            while (LineCount < RowNum)
            {
                string NameSpaceName = IXLW.Cell(++LineCount, 1).GetValue<string>();
                string TypeName = IXLW.Cell(++LineCount, 1).GetValue<string>();
                string MethodName = IXLW.Cell(++LineCount, 1).GetValue<string>();

                List<Instruction> InstructionList = GetInstructionList(LineCount, TypeDefList, NameSpaceName, TypeName, MethodName);

                if (InstructionList == null)
                {
                    while (IXLW.Cell(++LineCount, 1).GetValue<string>() != "--")
                    {
                        if (LineCount > RowNum)
                        {
                            break;
                        }
                    }

                    continue;
                }

                while (IXLW.Cell(++LineCount, 1).GetValue<string>() != "--")
                {
                    int Index = IXLW.Cell(LineCount, 1).GetValue<int>();
                    string ReadText = IXLW.Cell(LineCount, 2).GetValue<string>();
                    string WriteText = IXLW.Cell(LineCount, 3).GetValue<string>();

                    if (ReadText == WriteText)
                    {
                        continue;
                    }

                    Instruction Ins = null;

                    if (Index >= 0 && Index < InstructionList.Count)
                    {
                        Ins = InstructionList[Index];
                    }

                    if (Ins == null || Ins.Operand.ToString() != ReadText)
                    {
                        //Error
                        Console.WriteLine($"置き換えに失敗しました 行:[{LineCount}] NameSpace:[{NameSpaceName}] Type:[{TypeName}] Method:[{MethodName}] Text:[{WriteText}]\n");
                        ++ErrorCount;
                        continue;
                    }

                    Ins.Operand = WriteText;
                    ++SuccessCount;
                }
            }



            AssemblyDef.Write();

            Console.WriteLine("\n置き換えた文字列の数 : " + SuccessCount);
            Console.WriteLine("エラー発生回数 : " + ErrorCount);
            Console.WriteLine("\n文字列の置き換えが完了しました　キー入力で終了します");
            Console.ReadKey();
        }

        private static List<Instruction> GetInstructionList(int LineCount, IEnumerable<TypeDefinition> TypeDefList, string NameSpaceName, string TypeName, string MethodName)
        {
            //名前空間　検索
            IEnumerable<TypeDefinition> TargetNameSpace = TypeDefList.Where(x => x.Namespace == NameSpaceName);

            if (TargetNameSpace == null || TargetNameSpace.Count() == 0)
            {
                //Error
                Console.WriteLine($"ネームスペースが見つかりませんでした 行:[{LineCount}] NameSpace:[{NameSpaceName}]\n");
                ++ErrorCount;
                return null;
            }

            //クラス　検索
            TypeDefinition TargetType = TargetNameSpace.FirstOrDefault(x => x.Name == TypeName);

            if (TargetType == null)
            {
                //Error
                Console.WriteLine($"クラスが見つかりませんでした 行:[{LineCount}] NameSpace:[{NameSpaceName}] Type:[{TypeName}]\n");
                ++ErrorCount;
                return null;
            }

            //メソッド　検索
            MethodDefinition TargetMethod = null;

            foreach (MethodDefinition MD in TargetType.Methods)
            {
                string ParameterName = string.Join(",", MD.Parameters.Select(x => x.ParameterType.FullName));

                if ($"{MD.Name}({ParameterName})" == MethodName)
                {
                    TargetMethod = MD;
                }
            }

            if (TargetMethod == null || !TargetMethod.HasBody)
            {
                //Error
                Console.WriteLine($"メソッドが見つかりませんでした 行:[{LineCount}] NameSpace:[{NameSpaceName}] Type:[{TypeName}] Method:[{MethodName}]\n");
                ++ErrorCount;
                return null;
            }

            return TargetMethod.Body.Instructions.Where(x => x.OpCode == OpCodes.Ldstr).ToList();
        }
    }
}
