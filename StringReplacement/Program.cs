using ClosedXML.Excel;
using Mono.Cecil;
using Mono.Cecil.Cil;
using Newtonsoft.Json.Linq;
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
                Filter = "xlsx files (*.xlsx)|*.xlsx|json files (*.json)|*.json|All files (*.*)|*.*"
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

            string Extension = Path.GetExtension(InputPath);

            if (Extension == ".xlsx")
            {
                IXLWorksheet IXLW = new XLWorkbook(InputPath).Worksheet("Translation");
                int RowNum = IXLW.LastRowUsed().RowNumber();
                int Count0 = 0;

                while (Count0 <= RowNum)
                {
                    string NameSpaceName = IXLW.Cell(++Count0, 1).GetValue<string>();
                    string TypeName = IXLW.Cell(++Count0, 1).GetValue<string>();
                    string MethodName = IXLW.Cell(++Count0, 1).GetValue<string>();

                    List<Instruction> InstructionList = GetInstructionList(TypeDefList, NameSpaceName, TypeName, MethodName);

                    if (InstructionList == null)
                    {
                        continue;
                    }

                    while (IXLW.Cell(++Count0, 1).GetValue<string>() != "--")
                    {
                        int Index = IXLW.Cell(Count0, 1).GetValue<int>();
                        string ReadText = IXLW.Cell(Count0, 2).GetValue<string>();
                        string WriteText = IXLW.Cell(Count0, 3).GetValue<string>();
                        
                        Instruction Ins = null;

                        if (Index >= 0 && Index < InstructionList.Count)
                        {
                            Ins = InstructionList[Index];
                        }

                        if (Ins == null || Ins.Operand.ToString() != ReadText)
                        {
                            //Error
                            Console.WriteLine($"置き換えに失敗しました NameSpace:[{NameSpaceName}] Type:[{TypeName}] Method:[{MethodName}] Text:[{WriteText}]\n");
                            ++ErrorCount;
                            continue;
                        }

                        Ins.Operand = WriteText;
                        ++SuccessCount;
                    }
                }
            }
            else
            {
                JArray MainJObject = JArray.Parse(File.ReadAllText(InputPath));

                foreach (JArray ReplacementObject in MainJObject)
                {
                    string NameSpaceName = (string)ReplacementObject[0];
                    string TypeName = (string)ReplacementObject[1];
                    string MethodName = (string)ReplacementObject[2];

                    List<Instruction> InstructionList = GetInstructionList(TypeDefList, NameSpaceName, TypeName, MethodName);

                    if (InstructionList == null)
                    {
                        continue;
                    }

                    for (int Count0 = 3; Count0 < ReplacementObject.Count; ++Count0)
                    {
                        JArray TextObject = (JArray)ReplacementObject[Count0];

                        int Index = (int)TextObject[0];
                        string ReadText = (string)TextObject[1];
                        string WriteText = (string)TextObject[2];

                        Instruction Ins = InstructionList[Index];

                        if (Ins.Operand.ToString() != ReadText)
                        {
                            //Error
                            Console.WriteLine($"置き換えに失敗しました NameSpace:[{NameSpaceName}] Type:[{TypeName}] Method:[{MethodName}] Text:[{WriteText}]\n");
                            ++ErrorCount;
                            continue;
                        }

                        Ins.Operand = WriteText;
                        ++SuccessCount;
                    }
                }
            }

            AssemblyDef.Write();

            Console.WriteLine("\n置き換えた文字列の数 : " + SuccessCount);
            Console.WriteLine("エラー発生回数 : " + ErrorCount);
            Console.WriteLine("\n文字列の置き換えが完了しました　キー入力で終了します");
            Console.ReadKey();
        }

        private static List<Instruction> GetInstructionList(IEnumerable<TypeDefinition> TypeDefList, string NameSpaceName, string TypeName, string MethodName)
        {
            //名前空間　検索
            IEnumerable<TypeDefinition> TargetNameSpace = TypeDefList.Where(x => x.Namespace == NameSpaceName);

            if (TargetNameSpace == null || TargetNameSpace.Count() == 0)
            {
                //Error
                Console.WriteLine($"ネームスペースが見つかりませんでした NameSpace:[{NameSpaceName}]\n");
                ++ErrorCount;
                return null;
            }


            //クラス　検索
            TypeDefinition TargetType = TargetNameSpace.FirstOrDefault(x => x.Name == TypeName);

            if (TargetType == null)
            {
                //Error
                Console.WriteLine($"クラスが見つかりませんでした NameSpace:[{NameSpaceName}] Type:[{TypeName}]\n");
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
                Console.WriteLine($"メソッドが見つかりませんでした NameSpace:[{NameSpaceName}] Type:[{TypeName}] Method:[{MethodName}]\n");
                ++ErrorCount;
                return null;
            }

            return TargetMethod.Body.Instructions.Where(x => x.OpCode == OpCodes.Ldstr).ToList();
        }
    }
}
