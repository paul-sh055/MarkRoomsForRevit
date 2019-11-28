using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using System.Windows.Forms;
using System.Linq;



    [TransactionAttribute(TransactionMode.Manual)]
    public class LibColorRoom : IExternalCommand
    {
        Result IExternalCommand.Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
                //Получение объектов приложения и документа 
                UIApplication uiApp = commandData.Application;
                Document doc = uiApp.ActiveUIDocument.Document;

                //фильтруем по категории помещение
                FilteredElementCollector eSpase = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms);

                //получаем список элементов комнат квартир из общего количества помещений 
                List<Element> RoomList = new List<Element>();

                foreach (Element elem in eSpase)
                {
                    Parameter elemParam = elem.LookupParameter("ROM_Зона");
                    string sElemParam = elemParam.HasValue ? elemParam.AsString() : String.Empty;
                    Parameter levelParam2 = elem.LookupParameter("BS_Этаж");
                    string sLevelParam2 = levelParam2.HasValue ? levelParam2.AsString() : String.Empty;
            
                    if (sElemParam.Contains("Квартира") && sLevelParam2!="") { RoomList.Add(elem); }
                }
                    Element[] eRooms = RoomList.ToArray<Element>(); //преобразуем список в массив
                    int iRooms = eRooms.GetLength(0);


            if (iRooms > 0)
            { 
                //получаем список блоков (корпусов зданий)
                List<string> KorpList = new List<string>();
                    
                foreach (Element elem in eRooms)
                {
                    Parameter elemParam = elem.LookupParameter("BS_Блок");
                    string elemParamStr = elemParam.HasValue ? elemParam.AsString() : String.Empty;
                        
                    Parameter levelParam2 = elem.LookupParameter("BS_Этаж");
                    string sLevelParam2 = elemParam.HasValue ? elemParam.AsString() : String.Empty;


                    if (elemParamStr!="" && sLevelParam2!="") { KorpList.Add(elemParamStr); }
                }
                    KorpList.RemoveAll(t => string.IsNullOrEmpty(t)); //удаляем пустые строки
                    List<string> liIDs = KorpList.Distinct().ToList<string>(); //удаляем дубликаты
                    string[] KorpArr = liIDs.ToArray(); //преобразуем список в массив
                    int nKorp = KorpArr.GetLength(0);
            

                for (int b = 0; b < nKorp; b++) //корпусы зданий
                {
                    if (KorpArr[b] != "")
                    {
                        List<Element> RoomsKorpList = new List<Element>(); //объявляем список квартир блока
                        RoomsKorpList.Clear();

                        for (int r = 0; r < iRooms; r++)
                        {
                            if (eRooms[r] != null)
                            {
                                Parameter BS_Param = eRooms[r].LookupParameter("BS_Блок");
                                string BS_ParamStr = BS_Param.HasValue ? BS_Param.AsString() : String.Empty;
                                Parameter levelParam2 = eRooms[r].LookupParameter("BS_Этаж");
                                string sLevelParam2 = levelParam2.HasValue ? levelParam2.AsString() : String.Empty;

                                if (BS_ParamStr == KorpArr[b] && sLevelParam2!="") { RoomsKorpList.Add(eRooms[r]); }
                            }
                        }
                        
                        Element[] RoomsKorpArr = RoomsKorpList.ToArray(); //преобразуем список в массив    квартиры в блоке
                        int iRoomsKorp = RoomsKorpArr.GetLength(0);

                        List<string> LevelList = new List<string>(); //объявляем список этажей
                        string[] LevelArr1 = null;

                        List<string> RFlatList = new List<string>(); //объявляем список квартир ROM_Зона
                        string[] FlatArr1 = null; RFlatList.Clear();

                        List<string> TypeFlatList = new List<string>(); //объявляем список типов квартир ROM_Подзона
                        string[] TypeFlatArr1 = null; TypeFlatList.Clear();
                    
                        for (int r = 0; r < iRoomsKorp; r++)
                        {
                            if (RoomsKorpArr[r] != null)
                            {
                                Parameter levelparam = RoomsKorpArr[r].LookupParameter("BS_Этаж");
                                string levelparamStr = levelparam.HasValue ? levelparam.AsString() : String.Empty;
                                if (levelparamStr != "") { LevelList.Add(levelparamStr); }

                                Parameter eRFlatParam = RoomsKorpArr[r].LookupParameter("ROM_Зона");
                                string eRFlatParamStr = eRFlatParam.HasValue ? eRFlatParam.AsString() : String.Empty;
                                if (eRFlatParamStr != "") { RFlatList.Add(eRFlatParamStr); }

                                Parameter TypeParam = RoomsKorpArr[r].LookupParameter("ROM_Подзона");
                                string TypeParamStr = TypeParam.HasValue ? TypeParam.AsString() : String.Empty;
                                if (TypeParamStr != "") { TypeFlatList.Add(TypeParamStr); }

                            }
                        }

                        //список этажей
                        LevelList.RemoveAll(t => string.IsNullOrEmpty(t)); //удаляем пустые строки
                        List<string> liIDs1 = LevelList.Distinct().ToList<string>(); //удаляем дубликаты
                        LevelArr1 = liIDs1.ToArray(); //преобразуем список в массив
                        int iLevelCount = LevelArr1.GetLength(0);
                    
                        //список квартир
                        RFlatList.RemoveAll(t => string.IsNullOrEmpty(t)); //удаляем пустые строки
                        List<string> liIDs2 = RFlatList.Distinct().ToList<string>(); //удаляем дубликаты
                        FlatArr1 = liIDs2.ToArray(); //преобразуем список в массив
                        int iFlatCount = FlatArr1.GetLength(0);

                        //список типов квартир
                        TypeFlatList.RemoveAll(t => string.IsNullOrEmpty(t)); //удаляем пустые строки
                        List<string> liIDs3 = TypeFlatList.Distinct().ToList<string>(); //удаляем дубликаты
                        TypeFlatArr1 = liIDs3.ToArray(); //преобразуем список в массив
                        int iTypeCount = TypeFlatArr1.GetLength(0);
                    
                        Element[,,] eRoomlArr = new Element[iLevelCount + 1, iFlatCount + 1, 41];
                        string[,,] sRoomlArrL = new string[iLevelCount + 1, iFlatCount + 1, 41]; //массив помещений, параметр этаж
                        string[,,] sRoomlArrZ = new string[iLevelCount + 1, iFlatCount + 1, 41]; //массив помещений, параметр ROM_Зона только цифра
                        string[,,] s1RoomlArrZ = new string[iLevelCount + 1, iFlatCount + 1, 41]; //массив помещений, параметр ROM_Зона 
                        string[,,] sRoomlArrZP = new string[iLevelCount + 1, iFlatCount + 1, 41]; //массив помещений, параметр ROM_Подзона
                    

                        for (int lev = 0; lev < iLevelCount; lev++) //этажи
                        {
                            
                            if (LevelArr1[lev] != "")
                            { 
                                for (int fl = 0; fl < iFlatCount; fl++) //квартиры
                                {
                                    int r3 = 0;

                                    for (int r = 0; r < iRoomsKorp; r++) //квартирные помещения секции
                                    {
                                        if (RoomsKorpArr[r] != null)
                                        {
                                                    
                                            Parameter levelparam = RoomsKorpArr[r].LookupParameter("BS_Этаж");
                                            string levelparamStr = levelparam.HasValue ? levelparam.AsString() : String.Empty;

                                            Parameter Flat_Param = RoomsKorpArr[r].LookupParameter("ROM_Зона");
                                            string Flat_ParamStr = Flat_Param.HasValue ? Flat_Param.AsString() : String.Empty;

                                            Parameter TypeParam = RoomsKorpArr[r].LookupParameter("ROM_Подзона");
                                             string TypeParamStr = TypeParam.HasValue ? TypeParam.AsString() : String.Empty;

                                            if (levelparamStr != "" && Flat_ParamStr != "")
                                            { 
                                                if (levelparamStr == LevelArr1[lev] )
                                                {
                                                    if ((Flat_ParamStr == FlatArr1[fl]) && FlatArr1[fl] != "")
                                                    {
                                                        r3++; 
                                                        eRoomlArr[lev, fl, r3] = RoomsKorpArr[r];

                                                        sRoomlArrL[lev, fl, r3] = levelparamStr;

                                                        s1RoomlArrZ[lev, fl, r3] = Flat_ParamStr;
                                                        Flat_ParamStr = Flat_ParamStr.Replace(" ", "");
                                                        Flat_ParamStr = Flat_ParamStr.Replace("Квартира0", "");
                                                        Flat_ParamStr = Flat_ParamStr.Replace("Квартира", "") ;
                                                        sRoomlArrZ[lev, fl, r3] = Flat_ParamStr;

                                                        sRoomlArrZP[lev, fl, r3] = TypeParamStr;
                                                    
                                                    }
                                                }
                                            }
                                        }
                                    } 
                                } 
                            }
                        }

                        MessageBox.Show("Блок - " + KorpArr[b] + "\nЭтажей с квартирами - " + Convert.ToString(s1RoomlArrZ.GetLength(0)-1) + "\nКвартир в секции - " + Convert.ToString(s1RoomlArrZ.GetLength(1)-1) + "\nТипов квартир - " + Convert.ToString(iTypeCount));
                    

                        int iLEV = eRoomlArr.GetLength(0)-1; //этажей
                        int iFL = eRoomlArr.GetLength(1)-1;  //комнат
                        int iFLR = eRoomlArr.GetLength(2)-1; //помещений
                    

                       //составляем список элементов для маркировки

                        for (int lev = 0; lev < iLEV; lev++) //этажи
                        {
                                 
                            Element[,] eRoomlArr2 = new Element[iFL + 1, iFLR + 1];
                            string[,] sRoomlArrZ2 = new string [iFL + 1, iFLR + 1]; 
                            string[,] s1RoomlArrZ2 = new string[iFL + 1, iFLR + 1]; 
                            string[,] sRoomlArrZP2 = new string[iFL + 1, iFLR + 1];

                            for (int fl = 0; fl < iFL; fl++) //квартиры
                            {
                                for (int r = 1; r < iFLR; r++) //квартирные помещения
                                {
                                    eRoomlArr2[fl, r] = eRoomlArr[lev, fl, r];
                                    sRoomlArrZ2[fl, r] = sRoomlArrZ[lev, fl, r];
                                    s1RoomlArrZ2[fl, r] = s1RoomlArrZ[lev, fl, r];
                                    sRoomlArrZP2[fl, r] = sRoomlArrZP[lev, fl, r];
                                }
                            } 
                                  

                            int iFL2 = eRoomlArr2.GetLength(0)-1;  //комнат
                            int iFLR2 = eRoomlArr2.GetLength(1)-1; //помещений

                            List<string> sTypeFlatListDbl = new List<string>(); //список типов квартир с одинаковым количеством комнат
                                 
                            //список типов квартир, которые встречаются более 1 раза

                            for (int ty = 0; ty < iTypeCount; ty++)
                            {
                               int q = 0;

                                if (TypeFlatArr1[ty] != "")
                                {
                                    for (int fl = 0; fl < iFL2; fl++)
                                    {
                                        if (sRoomlArrZP2[fl, 1] == TypeFlatArr1[ty])
                                        {
                                            q++;
                                            if (q > 1)
                                            {
                                                sTypeFlatListDbl.Add(TypeFlatArr1[ty]);
                                            }
                                        }      
                                    }
                                }
                            }
                            
                            List<string> sTypeFlatDbl = sTypeFlatListDbl.Distinct().ToList<string>(); //удаляем дубликаты

                            string[] sTypeFlatArrDbl = sTypeFlatDbl.ToArray();
                            int iTypeFlatDbl = sTypeFlatArrDbl.GetLength(0);
                        
                            Element[,] eRoomlArr3 = new Element[iFL + 1, iFLR + 1];
                            string[,] sRoomlArrZ3 = new string [iFL + 1, iFLR + 1]; 
                            string[,] s1RoomlArrZ3 = new string[iFL + 1, iFLR + 1]; 
                            string[,] sRoomlArrZP3 = new string[iFL + 1, iFLR + 1];
                                 
                            for (int ty = 0; ty < sTypeFlatArrDbl.GetLength(0); ty++)
                            {
                                for (int fl = 0; fl < iFL2; fl++)
                                {
                                    if (sRoomlArrZP2[fl, 1] == sTypeFlatArrDbl[ty])
                                    {
                                        eRoomlArr3[fl, 1] = eRoomlArr2[fl, 1];
                                        sRoomlArrZ3[fl, 1] = sRoomlArrZ2[fl, 1];
                                        s1RoomlArrZ3[fl, 1] = s1RoomlArrZ2[fl, 1];
                                        sRoomlArrZP3[fl, 1] = sRoomlArrZP2[fl, 1];
                                    }      
                                }
                            }
                                 
                            //___________________________________
                               


                            //конвертируем в int номера квартир, записываем в массив int, расставляем номера квартир по возрастающей

                            int[] iRoomlArrZ = new int[iFL + 1]; 
                        
                            for (int i = 0; i < iFL2; i++)
                            {
                                int number;
                                bool result = int.TryParse (sRoomlArrZ3[i, 1], out number);
                                if (result == true) { iRoomlArrZ[i] = number; }
                                else{ iRoomlArrZ[i] = 0; }
                            }

                            List<int> iRoomList = new List<int>();
                            for (int i = 0; i < iFL2; i++) { if (iRoomlArrZ[i]!= 0) { iRoomList.Add(iRoomlArrZ[i]); } }
                            int[] iRoomlArrZ2 = iRoomList.ToArray();
                            Array.Sort(iRoomlArrZ2); iRoomlArrZ2.Distinct(); //сортируем номера квартир по возрастающей
                            int iNearFlatCount = iRoomlArrZ2.GetLength(0);
                        
                        

                            //сортируем массивы параметров квартир в соответствии со списком int по возрастанию

                            Element[,] eRoomlArr4 = new Element[iRoomlArrZ2.GetLength(0) + 1, iFLR + 1];
                            string[,] sRoomlArrZ4 = new string [iRoomlArrZ2.GetLength(0) + 1, iFLR + 1]; 
                            string[,] s1RoomlArrZ4 = new string[iRoomlArrZ2.GetLength(0) + 1, iFLR + 1]; 
                            string[,] sRoomlArrZP4 = new string[iRoomlArrZ2.GetLength(0) + 1, iFLR + 1];
                        
                            for (int f2 = 0; f2 < iRoomlArrZ2.GetLength(0); f2++)
                            {
                                for (int fl = 0; fl < iFL2; fl++)
                                {
                                    if (sRoomlArrZ3[fl, 1] == Convert.ToString(iRoomlArrZ2[f2]))
                                    {
                                        sRoomlArrZ4[f2, 1] = sRoomlArrZ3[fl, 1];
                                    }      
                                } 
                            }

                            //___________________________________

                            

                            //дописываем остальные комнаты к первой комнате
                            for (int fl = 0; fl < sRoomlArrZ4.GetLength(0); fl++)
                            {
                                for (int fl2 = 0; fl2 < eRoomlArr2.GetLength(0); fl2++)
                                {
                                    if (sRoomlArrZ4[fl, 1] == sRoomlArrZ2[fl2, 1])
                                    {
                                        for (int r = 1; r < eRoomlArr2.GetLength(1); r++)
                                        {
                                            eRoomlArr4[fl, r] = eRoomlArr2[fl2, r];
                                            sRoomlArrZ4[fl, r] = sRoomlArrZ2[fl2, r];
                                            s1RoomlArrZ4[fl, r] = s1RoomlArrZ2[fl2, r];
                                            sRoomlArrZP4[fl, r] = sRoomlArrZP2[fl2, r];
                                            
                                        }
                                    }
                                }
                            } 
                             



                            //если однотипная квартира следующая по номеру 
                            List<Element> eMarkRooms = new List<Element>(); eMarkRooms.Clear();

                            for (int fl = 0; fl < eRoomlArr4.GetLength(0); fl++)
                            {

                                //конвертируем в int номер первой квартиры
                                int numb1 = 0; int numbFL1 = 0;
                                string TypeFl1 = ""; string TypeFl2 = "";

                                TypeFl1 = sRoomlArrZP4[fl, 1];

                                bool result1 = int.TryParse (sRoomlArrZ4[fl, 1], out numb1);
                                if (result1 == true) { numbFL1 = numb1; }
                                else{ numbFL1 = 0; }


                                //конвертируем в int номер следующей по номеру квартиры
                                int numb2 = 0; int numbFL2 = 0;
                                if (fl < eRoomlArr4.GetLength(0)-1)
                                { 
                                    bool result2 = int.TryParse (sRoomlArrZ4[fl + 1, 1], out numb2);
                                    if (result2 == true) { numbFL2 = numb2; TypeFl2 = sRoomlArrZP4[fl + 1, 1]; }
                                    else{ numbFL2 = 0; TypeFl2 = ""; }
                                }
                                else
                                { numbFL2 = 0; TypeFl2 = ""; }



                                //если однотипная квартира следующая по номеру, добавляем в список
                                if (numbFL1 > 0 && numbFL2 > 0 && (numbFL1 + 1) == numbFL2)
                                {
                                    if (TypeFl1 == TypeFl2 && TypeFl1!="" && TypeFl1 != null)
                                    { 
                                        for (int r = 0; r < eRoomlArr4.GetLength(1); r++)
                                        {
                                            if (eRoomlArr4[fl, r]!= null)
                                            {
                                                eMarkRooms.Add(eRoomlArr4[fl, r]);
                                            }
                                        }
                                    }
                                }     
                            }
                                 


                            Element[] eMarkRoomsArr = eMarkRooms.ToArray();

                            for (int i = 0; i < eMarkRoomsArr.GetLength(0); i++)
                            {
                                if(eMarkRoomsArr[i]!=null)
                                {
                                    Parameter RomZID = eMarkRoomsArr[i].LookupParameter("ROM_Расчетная_подзона_ID");
                                    string sRomZID = RomZID.HasValue ? RomZID.AsString() : String.Empty;

                                    Parameter RomZIndex = eMarkRoomsArr[i].LookupParameter("ROM_Подзона_Index");
                                    string sRomZIndex = RomZIndex.HasValue ? RomZIndex.AsString() : String.Empty;

                                    Transaction transaction = new Transaction(doc, "Mark the rooms");
                                    transaction.Start();

                                    RomZIndex.Set(sRomZID + ".Полутон");

                                    transaction.Commit();
                                }
                            }

                                 


                        }
                        
    
                        DialogResult dialogResult = MessageBox.Show("Помещения блока " + KorpArr[b] + " промаркированы. Продолжить?", "Блок " + KorpArr[b], MessageBoxButtons.YesNo);
                        if(dialogResult == DialogResult.Yes)
                        {
                           //если нажали кнопку Да
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                           return Result.Cancelled;
                        }

                             

                    } 
                }
                    

            }
            else
            { MessageBox.Show("Помещения, с параметрами 'Квартира' на обнаружены"); }

           
        return Result.Succeeded;
        }
    }


