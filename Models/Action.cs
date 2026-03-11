using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text.Json;
using IPS_API.Services;
using IPS_API.Models.Data;

namespace IPS_API.Models.Commons
{
    public class IPS_Action
    {
        private readonly DB_Services _dB_Services;
        private readonly LogServices _logger;
        private readonly bool isProdEnv;
        private readonly DatabaseNames DB_1;
        private readonly DatabaseNames DB_2;

        private readonly string msMODULE_ID;
        public IPS_Action(DB_Services dB_Services, LogServices log_Services)
        {
            _dB_Services = dB_Services;
            isProdEnv = dB_Services.dB_Info.isProdEnv;
            DB_1 = dB_Services.dB_Info.DB_1;
            DB_2 = dB_Services.dB_Info.DB_2;
            _logger = log_Services;
            msMODULE_ID = "IPS_Action";
        }

        public DBstatus mCheckDB(DatabaseNames dbName, bool isProdEnv)
        {
            _logger.LogProcIn(msMODULE_ID, nameof(mCheckDB));
            bool isSuccess = false;
            try
            {
                string sql = "select count(*) CNT from v$session";
                var lst = DB_Services.Oracle.GetData<int>(dbName, isProdEnv, sql);
                isSuccess = lst.Any();
                if (isSuccess)
                {
                    sql = "select ora_database_name from dual";
                    var database_name_lst = DB_Services.Oracle.GetData<string>(dbName, isProdEnv, sql);
                }
            }
            catch (Exception ex)
            {
                isSuccess = false;
            }
            _logger.Append(msMODULE_ID, nameof(mCheckDB), "");

            _logger.LogProcOut(msMODULE_ID, nameof(mCheckDB));

            return new DBstatus { Success = isSuccess, ReturnMsg = "" };
        }

        private Random random = new Random((int)DateTime.Now.Ticks);
        private Dictionary<string, EQPInfo> EQP_Info = new Dictionary<string, EQPInfo>();
        private Dictionary<(string, string, string), MachineInfo> Machine_Info = new Dictionary<(string, string, string), MachineInfo>();
        private List<LotsInfo> Lots_Info = new List<LotsInfo>();
        private Dictionary<(string, string, string), List<LotsInfo>> ISSUE_NUMBER_Group = new Dictionary<(string, string, string), List<LotsInfo>>();
        private Dictionary<string, List<string>> ISSUE_NUMBERRoute = new Dictionary<string, List<string>>();
        #region 基因演算法主程式
        public (Chromosome, List<LotsInfo>) GA_Algorithm(GAInputData Input_Data, ref bool GADie)
        {
            string sProcID = nameof(GA_Algorithm);
            EQP_Info = new Dictionary<string, EQPInfo>();
            Machine_Info = new Dictionary<(string, string, string), MachineInfo>();
            Lots_Info = new List<LotsInfo>();
            ISSUE_NUMBER_Group = new Dictionary<(string, string, string), List<LotsInfo>>();
            Chromosome Best_Chromosome = new Chromosome();
            List<LotsInfo> Best_Lots_Temp = new List<LotsInfo>();
            ISSUE_NUMBERRoute = new Dictionary<string, List<string>>();
            string IPS_ID = Input_Data.IPS_ID;
            string CATEGORY = Input_Data.CATEGORY;
            double Crossover_Rate = Input_Data.Crossover_Rate;
            double Mutation_Rate  = Input_Data.Mutation_Rate;
            int Iteration_Times = Input_Data.Iteration_Times;
            int PopulationCnt = Input_Data.Population;
            double Retention_Ratio = Input_Data.Retention_Ratio;
            int ChangeSetupTime = Input_Data.ChangeSetupTime;
            // GA變數計算
            int RetentionNum = (int)(Retention_Ratio * PopulationCnt);
            int Crossover_RateNum = (int)(Crossover_Rate * PopulationCnt);
            int Mutation_RateNum = (int)(Mutation_Rate * PopulationCnt);
            #region 從DB撈取資料
            DataTable EQPInfoDt = null;
            DataTable MachineInfoDt = null;
            DataTable LotsInfoDt = null;
            string SQL = string.Empty;
            // EQPInfo
            SQL = $@"SELECT SQL";
            _logger.Append("準備要下SQL... " + SQL);
            EQPInfoDt = DB_Services.Oracle.GetDataTable(HDB, isProdEnv, SQL);
            // MachineInfo
            SQL = $@"SELECT SQL";
            _logger.Append("準備要下SQL... " + SQL);
            MachineInfoDt = DB_Services.Oracle.GetDataTable(HDB, isProdEnv, SQL);
            // LotsInfo
            SQL = $@"SELECT SQL";
            _logger.Append("準備要下SQL... " + SQL);
            LotsInfoDt = DB_Services.Oracle.GetDataTable(HDB, isProdEnv, SQL);
            #endregion

            #region 將DB資料寫入資料結構
            if (EQPInfoDt.Rows.Count > 0)
            {
                for (int i = 0; i < EQPInfoDt.Rows.Count; i++)
                {
                    DataRow row = EQPInfoDt.Rows[i];
                    EQPInfo Temp = new EQPInfo();
                    Temp.EQP = row["EQP"].ToString();
                    Temp.NOW_DEVICE = row["NOW_DEVICE"].ToString();
                    Temp.EQP_STATUS = row["EQP_STATUS"].ToString();
                    Temp.ETC = float.Parse(row["ETC"].ToString());
                    Temp.IPS_STEP = row["IPS_STEP"].ToString();                    
                    EQP_Info.Add(row["EQP"].ToString(), Temp);
                }
            }
            if (MachineInfoDt.Rows.Count > 0)
            {
                for (int i = 0; i < MachineInfoDt.Rows.Count; i++)
                {
                    DataRow row = MachineInfoDt.Rows[i];
                    MachineInfo Temp = new MachineInfo(); 
                    Temp.DEVICE_TYPE = row["DEVICE_TYPE"].ToString();                    
                    Temp.EQP = row["EQP"].ToString();
                    Temp.PREPARE_TIME = float.Parse(row["PREPARE_TIME"].ToString());
                    Temp.TO_END_STEP_TIME = float.Parse(row["TO_END_STEP_TIME"].ToString());
                    Temp.Priority = int.Parse(row["Priority"].ToString());
                    Temp.IPS_STEP = row["IPS_STEP"].ToString();
                    Temp.IPS_STEP_NEXT = row["IPS_STEP_NEXT"].ToString();
                    Temp.CATEGORY = row["CATEGORY"].ToString();
                    Machine_Info.Add((Temp.EQP, Temp.DEVICE_TYPE, Temp.IPS_STEP), Temp);
                }
            }
            if (LotsInfoDt.Rows.Count > 0)
            {
                for (int i = 0; i < LotsInfoDt.Rows.Count; i++)
                {
                    DataRow row = LotsInfoDt.Rows[i];
                    LotsInfo Temp = new LotsInfo();
                    Temp.IPS_ID = row["IPS_ID"].ToString();
                    Temp.LOT_ID = row["LOT_ID"].ToString();                    
                    Temp.IPS_STEP = row["IPS_STEP"].ToString();
                    Temp.Qty_pcs = float.Parse(row["Qty_pcs"].ToString());
                    Temp.DEVICE_TYPE = row["DEVICE_TYPE"].ToString();                    
                    Temp.WIPEND_TIME = row["WIPEND_TIME"].ToString();
                    Temp.Priority = int.Parse(row["Priority"].ToString());
                    Temp.ISSUE_NUMBER = row["ISSUE_NUMBER"].ToString();
                    Temp.CURRENT_STEP = row["CURRENT_STEP"].ToString();
                    Temp.TO_IPS_STEP_TIME = float.Parse(row["TO_IPS_STEP_TIME"].ToString());
                    Temp.TRACKIN_TIME = DateTime.Parse(Temp.ORIGINAL_TRACKIN_TIME).AddSeconds(Temp.TO_IPS_STEP_TIME).ToString("MM/dd/yyyy HH:mm:ss");
                    Temp.EQP = string.Empty;
                    Temp.TRACKOUT_TIME = string.Empty;
                    Temp.TO_END_STEP_TIME = 0;
                    Temp.ExceedWIPEND_TIME = false;
                    Temp.IS_CHANGE_EQP = false;
                    Temp.Memo = "";
                    Temp.CATEGORY = Input_Data.CATEGORY;
                    Lots_Info.Add(Temp);
                }
            }
            else
            {
                GADie = true;
                return (new Chromosome(), new List<LotsInfo>());
            }
            #endregion
            #region 將相同ISSUE_NUMBER之工作綁定在一起
            foreach (LotsInfo TempLotsInfo in Lots_Info)
            {
                if (!ISSUE_NUMBER_Group.ContainsKey((TempLotsInfo.ISSUE_NUMBER, TempLotsInfo.DEVICE_TYPE, TempLotsInfo.IPS_STEP)))
                {
                    ISSUE_NUMBER_Group[(TempLotsInfo.ISSUE_NUMBER, TempLotsInfo.DEVICE_TYPE, TempLotsInfo.IPS_STEP)] = new List<LotsInfo>();
                }
                ISSUE_NUMBER_Group[(TempLotsInfo.ISSUE_NUMBER, TempLotsInfo.DEVICE_TYPE, TempLotsInfo.IPS_STEP)].Add(TempLotsInfo);
            }
            foreach (var TempISSUE_NUMBERItem in ISSUE_NUMBER_Group)
            {
                TempISSUE_NUMBERItem.Value.Sort((a, b) => DateTime.Parse(a.WIPEND_TIME).Ticks.CompareTo(DateTime.Parse(b.WIPEND_TIME).Ticks));
                TempISSUE_NUMBERItem.Value.Sort((a, b) => DateTime.Parse(a.TRACKIN_TIME).Ticks.CompareTo(DateTime.Parse(b.TRACKIN_TIME).Ticks));
            }
            #endregion
            #region 產出相對應的ISSUE_NUMBER_Group

            foreach (var TempISSUE_NUMBERItem in ISSUE_NUMBER_Group)
            {
                string NowIPS_STEP = TempISSUE_NUMBERItem.Value[0].IPS_STEP;
                string NowProcessDEVICE_TYPE = TempISSUE_NUMBERItem.Value[0].DEVICE_TYPE;
                string NowISSUE_NUMBER = TempISSUE_NUMBERItem.Value[0].ISSUE_NUMBER;
                string IPS_STEP_NEXT = NowIPS_STEP;
                while (IPS_STEP_NEXT != "-1")
                {
                    bool Ck = false;
                    foreach (var TempMachineInfo in Machine_Info)
                    {
                        if (TempMachineInfo.Value.DEVICE_TYPE == NowProcessDEVICE_TYPE && TempMachineInfo.Value.IPS_STEP == IPS_STEP_NEXT)
                        {
                            if (!ISSUE_NUMBERRoute.ContainsKey(NowISSUE_NUMBER))
                            {
                                ISSUE_NUMBERRoute.Add(NowISSUE_NUMBER, new List<string>());
                            }
                            if (ISSUE_NUMBERRoute[NowISSUE_NUMBER].Exists(t => t == TempMachineInfo.Value.IPS_STEP_NEXT))
                            {
                                int index = ISSUE_NUMBERRoute[NowISSUE_NUMBER].IndexOf(TempMachineInfo.Value.IPS_STEP_NEXT);
                                ISSUE_NUMBERRoute[NowISSUE_NUMBER].Insert(index, TempMachineInfo.Value.IPS_STEP);
                                IPS_STEP_NEXT = TempMachineInfo.Value.IPS_STEP_NEXT;
                                Ck = true;
                                break;
                            }
                            else if (ISSUE_NUMBERRoute[NowISSUE_NUMBER].Exists(t => t == TempMachineInfo.Value.IPS_STEP))
                            {
                                Ck = true;
                                IPS_STEP_NEXT = TempMachineInfo.Value.IPS_STEP_NEXT;
                                break;
                            }
                            else
                            {
                                ISSUE_NUMBERRoute[NowISSUE_NUMBER].Add(TempMachineInfo.Value.IPS_STEP);
                                IPS_STEP_NEXT = TempMachineInfo.Value.IPS_STEP_NEXT;
                                Ck = true;
                                break;
                            }
                        }
                    }
                    if (Ck == false)
                    {
                        break;
                    }
                }
                if (GADie) break;
            }
            Dictionary<(string, string, string), List<LotsInfo>> TempISSUE_NUMBER_Group = new Dictionary<(string, string, string), List<LotsInfo>>();
            foreach (var TempISSUE_NUMBERItem in ISSUE_NUMBER_Group)
            {
                string NowISSUE_NUMBER = TempISSUE_NUMBERItem.Key.Item1;
                string NOW_DEVICE_TYPE = TempISSUE_NUMBERItem.Key.Item2;
                string NowIPS_STEP = TempISSUE_NUMBERItem.Key.Item3;
                if (ISSUE_NUMBERRoute.ContainsKey(NowISSUE_NUMBER) == false)
                {
                    continue;
                }
                int index = ISSUE_NUMBERRoute[NowISSUE_NUMBER].IndexOf(NowIPS_STEP);
                for (int i = index + 1; i < ISSUE_NUMBERRoute[NowISSUE_NUMBER].Count; i++)
                {
                    List<LotsInfo> TempProcessLotInfo = new List<LotsInfo>();
                    foreach (var Temp in TempISSUE_NUMBERItem.Value)
                    {
                        LotsInfo NextProcess = Temp.DeepCopy();
                        NextProcess.IPS_STEP = ISSUE_NUMBERRoute[NowISSUE_NUMBER][i];
                        TempProcessLotInfo.Add(NextProcess.DeepCopy());
                    }
                    if (!TempISSUE_NUMBER_Group.ContainsKey((NowISSUE_NUMBER, NOW_DEVICE_TYPE, ISSUE_NUMBERRoute[NowISSUE_NUMBER][i])))
                    {
                        TempISSUE_NUMBER_Group.Add((NowISSUE_NUMBER, NOW_DEVICE_TYPE, ISSUE_NUMBERRoute[NowISSUE_NUMBER][i]), new List<LotsInfo>());
                    }
                    TempISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, ISSUE_NUMBERRoute[NowISSUE_NUMBER][i])] = TempProcessLotInfo;
                }
            }
            foreach (var Temp in TempISSUE_NUMBER_Group)
            {
                if (!ISSUE_NUMBER_Group.ContainsKey(Temp.Key))
                {
                    ISSUE_NUMBER_Group.Add(Temp.Key, new List<LotsInfo>());
                }
                foreach (LotsInfo TempLotsInfo in Temp.Value)
                {
                    ISSUE_NUMBER_Group[Temp.Key].Add(TempLotsInfo.DeepCopy());
                }
            }
            #endregion

            #region 基因演算法計算開始
            Population CurrentPopulation = new Population(PopulationCnt, ISSUE_NUMBER_Group.Count);
            for (int i = 0; i < PopulationCnt; i++)
            {
                foreach (var TempItem in ISSUE_NUMBER_Group)
                {
                    CurrentPopulation.Chromosomes[i].ISSUE_NUMBERIndex.Add(TempItem.Key);
                }
            }
            float Fitness_Temp = 0.0f;
            int EXCEED_WIPEND_TIMES_COUNT_Temp = int.MaxValue;
            int ChangeSetupTimeS_COUNT_Temp = int.MaxValue;
            for (int i = 0; i < Iteration_Times; i++)
            {
                Population NewCurrentPopulation = new Population(PopulationCnt, Lots_Info.Count);
                NewCurrentPopulation.Chromosomes = new List<Chromosome>();
                #region GA-優勝劣汰
                for (int j = 0; j < PopulationCnt; j++)
                {
                    NewCurrentPopulation.Chromosomes.Add(CurrentPopulation.Chromosomes[j].DeepCopy());
                    NewCurrentPopulation.Chromosomes[j].ChangeSetupTimeS_COUNT = 0;
                    NewCurrentPopulation.Chromosomes[j].EXCEED_WIPEND_TIMES_COUNT = 0;
                    NewCurrentPopulation.Chromosomes[j].AllJobEndTime = 0;
                }
                CurrentPopulation = NewCurrentPopulation;
                #endregion
                #region GA-交配
                for (int j = 0; j < Crossover_RateNum; j++)
                {
                    int Parent1 = random.Next(RetentionNum, CurrentPopulation.Chromosomes.Count);
                    int Parent2 = random.Next(RetentionNum, CurrentPopulation.Chromosomes.Count);
                    CurrentPopulation.Chromosomes.Add(Crossover(CurrentPopulation.Chromosomes[Parent1], CurrentPopulation.Chromosomes[Parent2]));
                    CurrentPopulation.Chromosomes.Add(Crossover(CurrentPopulation.Chromosomes[Parent2], CurrentPopulation.Chromosomes[Parent1]));
                }
                #endregion
                #region GA-突變
                for (int j = 0; j < Mutation_RateNum && RetentionNum + j < CurrentPopulation.Chromosomes.Count; j++)
                {
                    CurrentPopulation.Chromosomes.Add(Mutate(CurrentPopulation.Chromosomes[RetentionNum + j]));
                }
                #endregion
                #region 遍歷每一條染色體計算
                for (int j = 0; j < CurrentPopulation.Chromosomes.Count; j++)
                {
                    CurrentPopulation.Chromosomes[j].SortChromosome(CurrentPopulation.Chromosomes[j].Genes.Count);
                    Dictionary<(string, string, string), MachineInfo> New_Machine_Info = new Dictionary<(string, string, string), MachineInfo>();
                    foreach (var Temp in Machine_Info)
                    {
                        if (!New_Machine_Info.ContainsKey(Temp.Key))
                            New_Machine_Info[Temp.Key] = Temp.Value.DeepCopy();
                    }
                    List<LotsInfo> New_Lots_Info = new List<LotsInfo>();
                    foreach (var Temp in Lots_Info)
                    {
                        New_Lots_Info.Add(Temp.DeepCopy());
                    }
                    Dictionary<string, EQPInfo> New_EQP_Info = new Dictionary<string, EQPInfo>();
                    foreach (var Temp in EQP_Info)
                    {
                        New_EQP_Info.Add(Temp.Key, Temp.Value.DeepCopy());
                    }
                    Dictionary<(string, string, string), List<LotsInfo>> New_ISSUE_NUMBER_Group = new Dictionary<(string, string, string), List<LotsInfo>>();
                    foreach (var Temp in ISSUE_NUMBER_Group)
                    {
                        New_ISSUE_NUMBER_Group.Add(Temp.Key, new List<LotsInfo>());
                        foreach (LotsInfo TempLotsInfo in Temp.Value)
                        {
                            New_ISSUE_NUMBER_Group[Temp.Key].Add(TempLotsInfo.DeepCopy());
                        }
                    }
                    CheckISSUE_NUMBEROrder(CurrentPopulation.Chromosomes[j], ISSUE_NUMBERRoute);
                    Dictionary<(string, string, string), List<LotsInfo>> CalcResult = new Dictionary<(string, string, string), List<LotsInfo>>();

                    try
                    {
                        CalcResult = CalculateAllJobTimes(CurrentPopulation.Chromosomes[j], New_Lots_Info, New_Machine_Info, New_EQP_Info, New_ISSUE_NUMBER_Group, ChangeSetupTime, ref GADie);
                    }
                    catch (Exception ex)
                    {
                        Console.Write("CalculateAllJobTimes出現異常:" + ex);
                    }
                    if (GADie)
                    {
                        break;
                    }
                    List<LotsInfo> CacultLotInfo = new List<LotsInfo>();
                    foreach (List<LotsInfo> TempLotsInfo in CalcResult.Values)
                    {
                        foreach (LotsInfo TempLot in TempLotsInfo)
                        {
                            CacultLotInfo.Add(TempLot.DeepCopy());
                        }
                    }
                    CurrentPopulation.Chromosomes[j].Fitness = CalculateFitness(CurrentPopulation.Chromosomes[j], CacultLotInfo, Input_Data);
                    if (
                       CurrentPopulation.Chromosomes[j].EXCEED_WIPEND_TIMES_COUNT < EXCEED_WIPEND_TIMES_COUNT_Temp
                        
                        ||
                            CurrentPopulation.Chromosomes[j].EXCEED_WIPEND_TIMES_COUNT == EXCEED_WIPEND_TIMES_COUNT_Temp &&
                            CurrentPopulation.Chromosomes[j].Fitness > Fitness_Temp

                    )
                    {
                        ChangeSetupTimeS_COUNT_Temp = CurrentPopulation.Chromosomes[j].ChangeSetupTimeS_COUNT;
                        EXCEED_WIPEND_TIMES_COUNT_Temp = CurrentPopulation.Chromosomes[j].EXCEED_WIPEND_TIMES_COUNT;
                        Fitness_Temp = CurrentPopulation.Chromosomes[j].Fitness;
                        Best_Chromosome = CurrentPopulation.Chromosomes[j].DeepCopy();
                        Best_Lots_Temp = new List<LotsInfo>();
                        foreach (var Item in CalcResult)
                        {
                            foreach (LotsInfo ItemLot in Item.Value)
                            {
                                Best_Lots_Temp.Add(ItemLot.DeepCopy());
                            }
                        }
                    }
                }
                if (GADie)
                {
                    break;
                }
                CurrentPopulation.Chromosomes.Sort((a, b) => b.Fitness.CompareTo(a.Fitness));
                Lots_Info = new List<LotsInfo>();
                foreach (LotsInfo Item in Best_Lots_Temp)
                {
                    Lots_Info.Add(Item);
                }
                #endregion
            }
            if (GADie)
                return new(new Chromosome(), new List<LotsInfo>());
            #endregion
            return (Best_Chromosome, Lots_Info);
        }
        #endregion

        #region 將計算結果存回DB
        public void SaveResultFiles(GAInputData Input_Data, (double, Chromosome, List<LotsInfo>) Output, string FileSavePath, string FileSaveName, bool IsOutPutFile, bool IsDrawPng)
        {
            int TOTOAL_EQP = 0;
            int TOTAL_OPEN_EQP = 0;
            StringBuilder EQPSQL = new StringBuilder();
            EQPSQL.Append("SELECT SQL");
            _logger.Append("準備要下SQL... " + EQPSQL);
            DataTable Dt = DB_Services.Oracle.GetDataTable(HDB, isProdEnv, EQPSQL.ToString());
            if (Dt.Rows.Count > 0)
            {
                TOTOAL_EQP = int.Parse(Dt.Rows[0]["TOTOAL_EQP"].ToString());
                TOTAL_OPEN_EQP = int.Parse(Dt.Rows[0]["TOTAL_OPEN_EQP"].ToString());
            }
            #region 將結果寫回DB
            string SQL = string.Empty;
            SQL += " INSERT SQL";
            _logger.Append("準備要下SQL... " + SQL);
            DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, SQL);
            for (int i = 0; i < Output.Item3.Count; i++)
            {
                SQL = string.Empty;
                SQL += " INSERT SQL ";
                _logger.Append("準備要下SQL... " + SQL);
                DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, SQL);
            }
            #endregion
        }
        #endregion

        #region 基因演算法要素：交配、突變
        public Chromosome Crossover(Chromosome parent1, Chromosome parent2)
        {
            Chromosome Result = new Chromosome();
            Result = parent1.DeepCopy();
            int CrossPoint1 = random.Next(0, parent1.Genes.Count);
            int CrossPoint2 = random.Next(CrossPoint1, parent1.Genes.Count);
            for (int i = CrossPoint1; i < CrossPoint2; i++)
            {
                Result.Genes[i] = parent2.Genes[i];
            }
            return Result;
        }
        public Chromosome Mutate(Chromosome chromosome)
        {
            Chromosome Result = new Chromosome();
            Result = chromosome.DeepCopy();
            Result.Genes[random.Next(chromosome.Genes.Count)] = (float)random.NextDouble();
            return Result;
        }
        #endregion

        #region 檢查每個ISSUE_NUMBER內的製程
        public void CheckISSUE_NUMBEROrder(Chromosome chromosome, Dictionary<string, List<string>> ISSUE_NUMBERRoute)
        {
            foreach (var item in ISSUE_NUMBERRoute)
            {
                string NowISSUE_NUMBER = item.Key;
                int TempCnt = 0;
                string NOW_DEVICE_TYPE = "";
                foreach (LotsInfo LotsItem in Lots_Info)
                {
                    if (LotsItem.ISSUE_NUMBER == NowISSUE_NUMBER && LotsItem.IPS_STEP == item.Value[0])
                    {
                        NOW_DEVICE_TYPE = LotsItem.DEVICE_TYPE;
                        break;
                    }
                }
                foreach (string RouteItem in item.Value)
                {
                    for (int i = TempCnt; i < chromosome.ISSUE_NUMBERIndex.Count; i++)
                    {
                        if (chromosome.ISSUE_NUMBERIndex[i].Item1 == NowISSUE_NUMBER)
                        {
                            chromosome.ISSUE_NUMBERIndex[i] = (NowISSUE_NUMBER, NOW_DEVICE_TYPE, RouteItem);
                            TempCnt = i + 1;
                            break;
                        }
                    }
                }
            }
        }
        #endregion

        #region 計算排程後的工時、基因演算法適應值
        public Dictionary<(string, string, string), List<LotsInfo>> CalculateAllJobTimes(Chromosome chromosome, List<LotsInfo> Lots_Info, Dictionary<(string, string, string), MachineInfo> Machine_Info, Dictionary<string, EQPInfo> EQP_Info, Dictionary<(string, string, string), List<LotsInfo>> ISSUE_NUMBER_Group, int ChangeSetupTime, ref bool GADie)
        {
            List<LotsInfo> Temp = new List<LotsInfo>();
            foreach (LotsInfo TempItem in Lots_Info)
                Temp.Add(TempItem.DeepCopy());
            Temp.Sort((a, b) => DateTime.Parse(a.ORIGINAL_TRACKIN_TIME).Ticks.CompareTo(DateTime.Parse(b.ORIGINAL_TRACKIN_TIME).Ticks));
            string FirstLotTRACKIN_TIME = Temp[0].ORIGINAL_TRACKIN_TIME; 
            float AllJobEndTime = 0;
            int ChangeSetupTimeS_COUNT = 0;
            int EXCEED_WIPEND_TIMES_COUNT = 0;
            for (int i = 0; i < chromosome.Genes.Count; i++) 
            {
                
                bool IsFirst = true;
                float FastStartTime = -1;
                string EQP_Index = "";
                bool NeedChangeSetup = false;
                int Machine_Priority = 1;
                string NowISSUE_NUMBER = chromosome.ISSUE_NUMBERIndex[i].Item1;
                string NOW_DEVICE_TYPE = chromosome.ISSUE_NUMBERIndex[i].Item2;
                string NowIPS_STEP = chromosome.ISSUE_NUMBERIndex[i].Item3;
                bool HasMac = false;
                bool HasOpenMac = false;
                try
                {
                    foreach (var MachineItem in Machine_Info)
                    {
                        if (!EQP_Info.ContainsKey(MachineItem.Value.EQP)) continue;
                        if (NOW_DEVICE_TYPE == MachineItem.Key.Item2 && NowIPS_STEP == MachineItem.Key.Item3)
                        {
                            HasMac = true;
                            int ChangeSetupTime_Gruel = 0;
                            float NowETC = EQP_Info[MachineItem.Value.EQP].ETC;
                            if ((int)NowETC == -1)
                            {
                                continue;
                            }
                            HasOpenMac = true;
                            if (!(EQP_Info[MachineItem.Value.EQP].NOW_DEVICE == NOW_DEVICE_TYPE && EQP_Info[MachineItem.Value.EQP].IPS_STEP == NowIPS_STEP))
                            
                            {
                                ChangeSetupTime_Gruel = ChangeSetupTime;
                            }
                            else
                                ChangeSetupTime_Gruel = 0;
                            if (IsFirst)
                            {
                                if (ChangeSetupTime_Gruel != 0)
                                    NeedChangeSetup = true;
                                else
                                    NeedChangeSetup = false;
                                FastStartTime = NowETC + ChangeSetupTime_Gruel;
                                EQP_Index = MachineItem.Key.Item1;
                                Machine_Priority = MachineItem.Value.Priority;
                                IsFirst = false;
                            }
                            else
                            {
                                if (NowETC + ChangeSetupTime_Gruel < FastStartTime ||
                                    (int)(NowETC + ChangeSetupTime_Gruel) == (int)FastStartTime && MachineItem.Value.Priority < Machine_Priority)
                                {
                                    if (ChangeSetupTime_Gruel != 0)
                                        NeedChangeSetup = true;
                                    else
                                        NeedChangeSetup = false;
                                    FastStartTime = NowETC + ChangeSetupTime_Gruel;
                                    EQP_Index = MachineItem.Key.Item1;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"DIE, {ex}.");
                }
                if (!HasMac)
                {
                    foreach (LotsInfo NowJobLot in ISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, NowIPS_STEP)])
                    {
                        NowJobLot.TRACKIN_TIME = "";
                        NowJobLot.TRACKOUT_TIME = "";
                        NowJobLot.EQP = "";
                        NowJobLot.Memo = "no machine";
                        NowJobLot.IS_CHANGE_EQP = false;
                        NowJobLot.ExceedWIPEND_TIME = false;
                    }
                }
                else if (HasMac && !HasOpenMac)
                {
                    foreach (LotsInfo NowJobLot in ISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, NowIPS_STEP)])
                    {
                        NowJobLot.TRACKIN_TIME = "";
                        NowJobLot.TRACKOUT_TIME = "";
                        NowJobLot.EQP = "";
                        NowJobLot.Memo = "no machine";
                        NowJobLot.IS_CHANGE_EQP = false;
                        NowJobLot.ExceedWIPEND_TIME = false;
                    }
                }
                else
                {
                    if (NeedChangeSetup)
                    {
                        ChangeSetupTimeS_COUNT++;
                        ISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, NowIPS_STEP)][0].IS_CHANGE_EQP = true;
                    }
                    else
                        ISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, NowIPS_STEP)][0].IS_CHANGE_EQP = false;
                    EQP_Info[EQP_Index].NOW_DEVICE = NOW_DEVICE_TYPE;
                    EQP_Info[EQP_Index].IPS_STEP = NowIPS_STEP;
                    float TO_END_STEP_TIME = Machine_Info[(EQP_Index, NOW_DEVICE_TYPE, NowIPS_STEP)].TO_END_STEP_TIME;
                    float NowMacPREPARE_TIME = Machine_Info[(EQP_Index, NOW_DEVICE_TYPE, NowIPS_STEP)].PREPARE_TIME;
                    if (FastStartTime < NowMacPREPARE_TIME) FastStartTime = NowMacPREPARE_TIME;
                    int Cnt = 0;
                    foreach (LotsInfo NowJobLot in ISSUE_NUMBER_Group[(NowISSUE_NUMBER, NOW_DEVICE_TYPE, NowIPS_STEP)])
                    {
                        NowJobLot.EQP = EQP_Index;
                        NowJobLot.PREPARE_TIME = Machine_Info[(EQP_Index, NOW_DEVICE_TYPE, NowIPS_STEP)].PREPARE_TIME;
                        NowJobLot.UPH_Pcs = Machine_Info[(EQP_Index, NOW_DEVICE_TYPE, NowIPS_STEP)].UPH_Pcs;
                        float AddSec = NowJobLot.Qty_pcs * Machine_Info[(EQP_Index, NOW_DEVICE_TYPE, NowIPS_STEP)].UPH_Pcs;
                        if (DateTime.Parse(FirstLotTRACKIN_TIME).AddSeconds(FastStartTime) < DateTime.Parse(NowJobLot.TRACKIN_TIME)) 
                        {
                            NowJobLot.TRACKIN_TIME = DateTime.Parse(NowJobLot.TRACKIN_TIME).ToString("MM/dd/yyyy HH:mm:ss");
                            FastStartTime = (float)(DateTime.Parse(NowJobLot.TRACKIN_TIME) - DateTime.Parse(FirstLotTRACKIN_TIME)).TotalSeconds;
                        }
                        else
                        {
                            NowJobLot.TRACKIN_TIME = DateTime.Parse(FirstLotTRACKIN_TIME).AddSeconds(FastStartTime).ToString("MM/dd/yyyy HH:mm:ss");
                        }
                        NowJobLot.TRACKOUT_TIME = DateTime.Parse(NowJobLot.TRACKIN_TIME).AddSeconds(AddSec).ToString("MM/dd/yyyy HH:mm:ss");
                        NowJobLot.TO_END_STEP_TIME = TO_END_STEP_TIME;
                        if (DateTime.Parse(NowJobLot.TRACKOUT_TIME).AddSeconds(TO_END_STEP_TIME) > DateTime.Parse(NowJobLot.WIPEND_TIME))
                        {  
                            EXCEED_WIPEND_TIMES_COUNT++;
                            NowJobLot.ExceedWIPEND_TIME = true;
                        }
                        else
                            NowJobLot.ExceedWIPEND_TIME = false;
                        float TempETC = FastStartTime + AddSec;
                        if (DateTime.Parse(FirstLotTRACKIN_TIME).AddSeconds(FastStartTime) < DateTime.Parse(NowJobLot.TRACKIN_TIME))
                        {
                            TempETC = (float)(DateTime.Parse(NowJobLot.TRACKOUT_TIME) - DateTime.Parse(FirstLotTRACKIN_TIME)).TotalSeconds;
                        }
                        EQP_Info[EQP_Index].ETC = TempETC;
                        if (TempETC > AllJobEndTime) AllJobEndTime = TempETC;
                        FastStartTime += AddSec;
                        bool Flag = false;

                        foreach (string LaterIssue in ISSUE_NUMBERRoute[NowISSUE_NUMBER])
                        {
                            if (Flag)
                            {
                                foreach (var NextISSUE_NUMBERGroup in ISSUE_NUMBER_Group)
                                {
                                    try
                                    {
                                        if (NextISSUE_NUMBERGroup.Key.Item1 == NowISSUE_NUMBER && NextISSUE_NUMBERGroup.Key.Item2 == NOW_DEVICE_TYPE && NextISSUE_NUMBERGroup.Key.Item3 == LaterIssue
                                            && NextISSUE_NUMBERGroup.Value[Cnt].LOT_ID == NowJobLot.LOT_ID)
                                        {

                                            NextISSUE_NUMBERGroup.Value[Cnt].TRACKIN_TIME = DateTime.Parse(NowJobLot.TRACKOUT_TIME).ToString("MM/dd/yyyy HH:mm:ss");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"{NowISSUE_NUMBER}製程({LaterIssue})有問題, {ex}.");
                                    }
                                }
                            }
                            if (LaterIssue == NowIPS_STEP) Flag = true;
                        }
                        Cnt += 1;
                    }
                }
            }
            if (GADie) return new Dictionary<(string, string, string), List<LotsInfo>>();
            chromosome.AllJobEndTime = AllJobEndTime;
            chromosome.ChangeSetupTimeS_COUNT = ChangeSetupTimeS_COUNT;
            chromosome.EXCEED_WIPEND_TIMES_COUNT = EXCEED_WIPEND_TIMES_COUNT;
            return ISSUE_NUMBER_Group;
        }
        public float CalculateFitness(Chromosome chromosome, List<LotsInfo> Lots_Info, GAInputData Ga_Info)
        {
            float fitness = 0;
            float AllJobEndTime_Gap = (float)Math.Pow(10, ((int)chromosome.AllJobEndTime * Ga_Info.AllJobEndTime_Weights).ToString().Length - 1);
            List<LotsInfo> TempLots_Info = new List<LotsInfo>();
            foreach (var Temp in Lots_Info)
            {
                if (Temp.EQP == "無機台可選")
                {

                }
                else
                    TempLots_Info.Add(Temp.DeepCopy());
            }
            float ExceedWIPEND_TIMESecs = 0;
            for (int i = 0; i < TempLots_Info.Count; i++)
            {
                if (TempLots_Info[i].ExceedWIPEND_TIME)
                {
                    float ExceepSeconds = (float)(DateTime.Parse(TempLots_Info[i].TRACKOUT_TIME).AddSeconds(TempLots_Info[i].TO_END_STEP_TIME) - DateTime.Parse(TempLots_Info[i].WIPEND_TIME)).TotalSeconds;
                    ExceedWIPEND_TIMESecs += ExceepSeconds * ExceepSeconds;
                }
            }
            ExceedWIPEND_TIMESecs = ExceedWIPEND_TIMESecs / TempLots_Info.Count;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            fitness++; fitness--;
            for (int i = 0; i < chromosome.Genes.Count; i++)
            {
                int LotsIndex = 0;
                for (LotsIndex = 0; LotsIndex < TempLots_Info.Count; LotsIndex++)
                {
                    if (TempLots_Info[LotsIndex].ISSUE_NUMBER == chromosome.ISSUE_NUMBERIndex[i].Item1 &&
                        TempLots_Info[LotsIndex].IPS_STEP == chromosome.ISSUE_NUMBERIndex[i].Item3)
                    {
                        break;
                    }
                }
                foreach (var Calculation in Ga_Info.Calculation_Formula)
                {
                    float Weights = Calculation.Weights * (chromosome.Genes.Count - i) / Genes_Count_Gap;
                    switch (Calculation.Variable_Name)
                    {
                        case "{Qty_pcs}":
                            break;
                        case "{WIPEND_TIME}":
                            break;
                        case "{LotsPriority}":
                            break;
                        case "{DEVICE_TYPE}":
                            break;
                        case "{LOT_ID}":
                            break;
                    }
                }
            }
            return fitness;
        }
        #endregion

        #region 
        public void DrawPng(string FileSavePath, string FileSaveName, (double, Chromosome, List<LotsInfo>) Result, string OriginTRACKIN_TIME, GAInputData inputData, string Method)
        {
            double ProcessTime = Result.Item1;
            Chromosome BestChromosome = Result.Item2;
            List<LotsInfo> Lots_Info = Result.Item3;
            DirectoryInfo Dir = new DirectoryInfo(FileSavePath);
            if (!Dir.Exists)
                Dir.Create();
            Dictionary<string, Color> ColorTable = new Dictionary<string, Color>();
            Bitmap bm = new Bitmap(1600, 700);
            Graphics g = Graphics.FromImage(bm);
            g.Clear(Color.FromArgb(200, 200, 200));
            List<string> EQP_Info = new List<string>();
            long StartTime = DateTime.Parse(OriginTRACKIN_TIME).Ticks / 10000000;
            long EndTime = -1;
            foreach (LotsInfo Lot in Lots_Info)
            {
                bool IsExist = EQP_Info.Exists(t => t == Lot.EQP);
                if (!IsExist)
                    EQP_Info.Add(Lot.EQP);
                long Temp = DateTime.Parse(Lot.TRACKOUT_TIME).Ticks / 10000000;
                if (EndTime == -1) EndTime = Temp;
                else if (EndTime < Temp) EndTime = Temp;
            } 
            EQP_Info.Sort();
            Brush DrawBrush = new SolidBrush(Color.Black);
            Font DrawFont = new Font("Arial", 10);
            for (int i = 0; i < EQP_Info.Count; i++)
                g.DrawString(EQP_Info[i], DrawFont, DrawBrush, 20, 25 * i);
            int Days = (int)((EndTime - StartTime) / 86400) + 1;
            for (int i = 0; i <= Days; i++)
            {
                DateTime TempStartDay = new DateTime(StartTime * 10000000);
                string Date = TempStartDay.AddSeconds(86400 * i).ToShortDateString();
                g.DrawString(Date, DrawFont, DrawBrush, 100 + 800 / Days * i, 650);
                g.DrawLine(new Pen(Color.Black, 2), 100 + 800 / Days * i, 0, 100 + 800 / Days * i, 700);
            }
            g.DrawRectangle(new Pen(Color.Black, 1), 1100, 20, 300, 800);
            
            g.DrawString("本次計算之IPS_ID：" + inputData.IPS_ID, new Font("Arial", 12), DrawBrush, 1100, 30);
            g.DrawString("交配比例：" + inputData.Crossover_Rate, DrawFont, DrawBrush, 1150, 70);
            g.DrawString("突變比例：" + inputData.Mutation_Rate, DrawFont, DrawBrush, 1150, 90);
            g.DrawString("迭代次數：" + inputData.Iteration_Times, DrawFont, DrawBrush, 1150, 110);
            g.DrawString("染色體數：" + inputData.Population, DrawFont, DrawBrush, 1150, 130);
            g.DrawString("擇優比例：" + inputData.Retention_Ratio, DrawFont, DrawBrush, 1150, 150);
            g.DrawString("工時權重：" + inputData.AllJobEndTime_Weights, DrawFont, DrawBrush, 1150, 170);
            g.DrawString("改機懲罰：" + inputData.ChangeSetupTime + "秒", DrawFont, DrawBrush, 1150, 190);
            g.DrawString("GA計算時間：" + ProcessTime.ToString() + "秒", DrawFont, new SolidBrush(Color.Red), 1150, 210);
            g.DrawString("改機次數：" + BestChromosome.ChangeSetupTimeS_COUNT + "次", DrawFont, new SolidBrush(Color.Red), 1150, 240);
            g.DrawString("超時數量：" + BestChromosome.EXCEED_WIPEND_TIMES_COUNT + "筆", DrawFont, new SolidBrush(Color.Red), 1150, 260);
            g.FillRectangle(new SolidBrush(Color.LightGreen), 1150, 280, 20, 20);
            g.DrawRectangle(new Pen(Color.Yellow, 1), 1150, 280, 20, 20);
            g.DrawString("換機時光", DrawFont, DrawBrush, 1200, 280);
            double Gap = (float)800 / (EndTime - StartTime);
            float ChangeEQPTime = inputData.ChangeSetupTime * (float)Gap;
            for (int i = 0; i < Lots_Info.Count; i++)
            {
                float X_Start = 0, X_End = 0, Y = 0;
                for (int j = 0; j < EQP_Info.Count; j++)
                {
                    if (EQP_Info[j] == Lots_Info[i].EQP)
                    {
                        Y = 25 * j;
                        break;
                    }
                }
                DateTime TempStartDay = new DateTime(StartTime * 10000000);
                string Date = TempStartDay.AddSeconds(86400 * i).ToShortDateString();
                X_Start = (float)(100 + (DateTime.Parse(Lots_Info[i].TRACKIN_TIME).Ticks / 10000000 - StartTime) * Gap);
                X_End = (float)(100 + (DateTime.Parse(Lots_Info[i].TRACKOUT_TIME).Ticks / 10000000 - StartTime) * Gap);
                Random rnd = new Random();
                Color color;
                if (!ColorTable.TryGetValue(Lots_Info[i].DEVICE_TYPE, out color))
                {
                    ColorTable[Lots_Info[i].DEVICE_TYPE] = Color.FromArgb(rnd.Next() % 255, rnd.Next(1, 150) % 255, rnd.Next() % 255);
                }
                color = ColorTable[Lots_Info[i].DEVICE_TYPE];
                Brush DrawRectangleBrush = new SolidBrush(color);
                if (Lots_Info[i].IS_CHANGE_EQP == true) 
                {
                    g.FillRectangle(new SolidBrush(Color.LightGreen), X_Start - ChangeEQPTime, Y, ChangeEQPTime, 15);
                    g.DrawRectangle(new Pen(Color.Yellow, 0.5f), X_Start - ChangeEQPTime, Y, ChangeEQPTime, 15);
                }
                g.FillRectangle(DrawRectangleBrush, X_Start, Y, X_End - X_Start, 15);
                if (Lots_Info[i].ExceedWIPEND_TIME == true)
                    g.DrawRectangle(new Pen(Color.Red, 0.5f), X_Start, Y, X_End - X_Start, 15);
                else
                    g.DrawRectangle(new Pen(Color.Black, 0.5f), X_Start, Y, X_End - X_Start, 15);
            }
            int Temp2 = 1;
            foreach (var TempColorInfo in ColorTable)
            {
                Color TempColor = TempColorInfo.Value;
                g.FillRectangle(new SolidBrush(TempColor), 1150, 290 + 20 * Temp2, 20, 20);
                g.DrawRectangle(new Pen(Color.Black, 1), 1150, 290 + 20 * Temp2, 20, 20);
                g.DrawString(TempColorInfo.Key, DrawFont, DrawBrush, 1200, 290 + 20 * Temp2);
                Temp2 += 1;
            }
            g.DrawLine(new Pen(Color.Blue, 3), 100, 0, 100, 900);
            FileSavePath = FileSavePath + FileSaveName + Method + "(Security C).png";
            bm.Save(FileSavePath);
            bm.Dispose();
            g.Dispose();
        }
        public void OutputCSV(string FileSavePath, string FileSaveName, List<LotsInfo> Result, string Method)
        {
            DirectoryInfo Dir = new DirectoryInfo(FileSavePath);
            if (!Dir.Exists)
                Dir.Create();
            FileSavePath = FileSavePath + FileSaveName + Method + "(Security C).csv";
            StringBuilder output = new StringBuilder();
            string[] headings = { "TT_LOT_ID", "Device_Type", "製程", "開始作業時間", "完成作業時間", "EQP" };
            output.AppendLine(string.Join(",", headings));
            foreach (var Temp in Result)
            {
                string[] newLine = { Temp.LOT_ID, Temp.DEVICE_TYPE, Temp.IPS_STEP, Temp.TRACKIN_TIME, Temp.TRACKOUT_TIME, Temp.EQP };
                output.AppendLine(string.Join(",", newLine));
            }
            try
            {
                File.AppendAllText(FileSavePath, output.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Data could not be written to the CSV file, {ex}.");
            }
        }
        public void OutputJson(string FileSavePath, string FileSaveName, (double, Chromosome, List<LotsInfo>) Result, string Method)
        {
            DirectoryInfo Dir = new DirectoryInfo(FileSavePath);
            if (!Dir.Exists)
                Dir.Create();
            FileSavePath = FileSavePath + FileSaveName + Method + "(Security C).json";
            File.WriteAllText(FileSavePath, JsonConvert.SerializeObject(Result, Formatting.Indented));
        }
        public void DrawHtml(string FileSavePath, string FileSaveName, (double, Chromosome, List<LotsInfo>) Output, string OriginTRACKIN_TIME, GAInputData inputData, string Method)
        {
            double ProcessTime = Output.Item1;
            Chromosome BestChromosome = Output.Item2;
            List<LotsInfo> Lots_Info = Output.Item3;
            DirectoryInfo Dir = new DirectoryInfo(FileSavePath);
            if (!Dir.Exists)
            {
                Dir.Create();
            }
            Dictionary<string, Color> ColorTable = new Dictionary<string, Color>();
            Color color;
            Random rnd = new Random();
            StreamWriter streamWriter = new StreamWriter(FileSavePath + FileSaveName + Method + "(Security C).html");
            streamWriter.WriteLine("<html>");
            streamWriter.WriteLine("    <head>");
            streamWriter.WriteLine("      <link href = 'https://cdn.jsdelivr.net/npm/bootstrap@5.0.0/dist/css/bootstrap.min.css' rel = 'stylesheet'");
            streamWriter.WriteLine("        integrity = 'sha384-wEmeIV1mKuiNpC+IOBjI7aAzPcEZeedi5yW5f2yOq55WWLwNGmvvx4Um1vskeMj0' crossorigin = 'anonymous'>");
            streamWriter.WriteLine("    </head>");
            streamWriter.WriteLine("    <body>");
            streamWriter.WriteLine("      <div>");
            streamWriter.WriteLine("        <table class='table table-hover' style='width:50%;'>");
            streamWriter.WriteLine("          <tr>");
            streamWriter.WriteLine("            <th colspan = '7' style='text-align:center'>基因演算法參數</th>");
            streamWriter.WriteLine("          </tr>");
            streamWriter.WriteLine("          <tr>");
            streamWriter.WriteLine("            <td colspan = '7' style='text-align:center'>本次計算之IPS_ID：" + inputData.IPS_ID + "</td>");
            streamWriter.WriteLine("          </tr>");
            streamWriter.WriteLine("          <tr>");
            streamWriter.WriteLine("            <td>交配比例</td>");
            streamWriter.WriteLine("            <td>突變比例</td>");
            streamWriter.WriteLine("            <td>迭代次數</td>");
            streamWriter.WriteLine("            <td>染色體數</td>");
            streamWriter.WriteLine("            <td>擇優比例</td>");
            streamWriter.WriteLine("            <td>工時權重</td>");
            streamWriter.WriteLine("            <td>改機處罰</td>");
            streamWriter.WriteLine("          </tr> <tr>");
            streamWriter.WriteLine("            <td>" + inputData.Crossover_Rate + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.Mutation_Rate + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.Iteration_Times + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.Population + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.Retention_Ratio + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.AllJobEndTime_Weights + "</td>");
            streamWriter.WriteLine("            <td>" + inputData.ChangeSetupTime + "秒</td>");
            streamWriter.WriteLine("          </tr>");
            streamWriter.WriteLine("          <tr>");
            streamWriter.WriteLine("            <td colspan='3'>基因演算法計算時間</td>");
            streamWriter.WriteLine("            <td colspan='2'>改機次數</td>");
            streamWriter.WriteLine("            <td colspan='2'>超時數量</td>");
            streamWriter.WriteLine("          </tr> <tr>");
            streamWriter.WriteLine("            <td colspan='3'style='color:red'>" + ProcessTime + "秒</td>");
            streamWriter.WriteLine("            <td colspan='2'style='color:red'>" + BestChromosome.ChangeSetupTimeS_COUNT + "次</td>");
            streamWriter.WriteLine("            <td colspan='2'style='color:red'>" + BestChromosome.EXCEED_WIPEND_TIMES_COUNT + "筆</td>");
            streamWriter.WriteLine("          </tr>");
            streamWriter.WriteLine("        </table>");
            streamWriter.WriteLine("      </div>");
            streamWriter.WriteLine("<script type = 'text/javascript' src = 'https://www.gstatic.com/charts/loader.js'></script>");
            streamWriter.WriteLine("<script type = 'text/javascript'>");
            streamWriter.WriteLine("  google.charts.load('current', { packages: ['timeline']});");
            streamWriter.WriteLine("            google.charts.setOnLoadCallback(drawChart);");
            streamWriter.WriteLine("            function drawChart()");
            streamWriter.WriteLine("            {");
            streamWriter.WriteLine("                var container = document.getElementById('GAChart');");
            streamWriter.WriteLine("                var chart = new google.visualization.Timeline(container);");
            streamWriter.WriteLine("                var dataTable = new google.visualization.DataTable();");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'string', id: 'EQP' });");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'string', id: 'DEVICE_TYPE' });");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'string', role: 'tooltip' });");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'string', id: 'style', role: 'style' });");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'date', id: 'Start' });");
            streamWriter.WriteLine("            dataTable.addColumn({ type: 'date', id: 'End' });");
            streamWriter.WriteLine("            let blankStyle = 'opacity: 0';");
            streamWriter.WriteLine("            let ExceedTimeStyle = 'stroke-width: 5;stroke-color: #ff0000;fill-color: #FFC0CB'");
            streamWriter.WriteLine("            dataTable.addRows([");
            List<LotsInfo> TempLotsInfo = new List<LotsInfo>();
            foreach (LotsInfo Temp in Lots_Info)
            {
                TempLotsInfo.Add(Temp.DeepCopy());
            }
            TempLotsInfo.Sort((a, b) => DateTime.Parse(a.TRACKIN_TIME).Ticks.CompareTo(DateTime.Parse(b.TRACKIN_TIME).Ticks));
            foreach (var TempEQP in EQP_Info)
            {
                DateTime TRACKIN_TIME = Convert.ToDateTime(TempLotsInfo[0].TRACKIN_TIME);
                streamWriter.WriteLine("      ['" + TempEQP.Value.EQP + "', '','',blankStyle,new Date(" + TRACKIN_TIME.Year + "," + (TRACKIN_TIME.Month - 1) + "," + TRACKIN_TIME.Day + "," + TRACKIN_TIME.Hour + "," + TRACKIN_TIME.Minute + "," + TRACKIN_TIME.Second + "," + "),new Date(" + TRACKIN_TIME.Year + "," + (TRACKIN_TIME.Month - 1) + "," + TRACKIN_TIME.Day + "," + TRACKIN_TIME.Hour + "," + TRACKIN_TIME.Minute + "," + TRACKIN_TIME.Second + "," + ")],");
            }
            for (int i = 0; i < Lots_Info.Count; i++)
            {
                if (!ColorTable.TryGetValue(Lots_Info[i].DEVICE_TYPE, out color))
                {
                    ColorTable[Lots_Info[i].DEVICE_TYPE] = Color.FromArgb(rnd.Next(1, 150) % 255, rnd.Next() % 255, rnd.Next() % 255);
                }
                color = ColorTable[Lots_Info[i].DEVICE_TYPE];
                DateTime TRACKIN_TIME = Convert.ToDateTime(Lots_Info[i].TRACKIN_TIME);
                DateTime TRACKOUT_TIME = Convert.ToDateTime(Lots_Info[i].TRACKOUT_TIME);
                if (Lots_Info[i].ExceedWIPEND_TIME)
                    streamWriter.WriteLine("      ['" + Lots_Info[i].EQP + "', '','<font color=#FF0000;>批號：" + Lots_Info[i].LOT_ID + "<br>DEVICE_TYPE：" + Lots_Info[i].DEVICE_TYPE + "<br>製程階段：" + Lots_Info[i].IPS_STEP + "<br>ISSUE_NUMBER：" + Lots_Info[i].ISSUE_NUMBER + "<br>-----<br>最快到站時間：" + DateTime.Parse(Lots_Info[i].ORIGINAL_TRACKIN_TIME).AddSeconds(Lots_Info[i].TO_IPS_STEP_TIME).ToString() + "<br>實際TrackIn時間：" + Lots_Info[i].TRACKIN_TIME + "<br>TrackOut時間：" + Lots_Info[i].TRACKOUT_TIME + "<br>To5400Time：" + DateTime.Parse(Lots_Info[i].TRACKOUT_TIME).AddSeconds((double)Lots_Info[i].TO_END_STEP_TIME) + "<br>期望交期：" + DateTime.Parse(Lots_Info[i].WIPEND_TIME).ToString() + "</font>',ExceedTimeStyle, new Date(" + TRACKIN_TIME.Year + "," + (TRACKIN_TIME.Month - 1) + "," + TRACKIN_TIME.Day + "," + TRACKIN_TIME.Hour + "," + TRACKIN_TIME.Minute + "," + TRACKIN_TIME.Second + "," + "), new Date(" + TRACKOUT_TIME.Year + "," + (TRACKOUT_TIME.Month - 1) + "," + TRACKOUT_TIME.Day + "," + TRACKOUT_TIME.Hour + "," + TRACKOUT_TIME.Minute + "," + TRACKOUT_TIME.Second + "," + ")],");
                else
                    streamWriter.WriteLine("      ['" + Lots_Info[i].EQP + "', '','批號：" + Lots_Info[i].LOT_ID + "<br>DEVICE_TYPE：" + Lots_Info[i].DEVICE_TYPE + "<br>製程階段：" + Lots_Info[i].IPS_STEP + "<br>ISSUE_NUMBER：" + Lots_Info[i].ISSUE_NUMBER + "<br>-----<br>最快到站時間：" + DateTime.Parse(Lots_Info[i].ORIGINAL_TRACKIN_TIME).AddSeconds(Lots_Info[i].TO_IPS_STEP_TIME).ToString() + "<br>實際TrackIn時間：" + Lots_Info[i].TRACKIN_TIME + "<br>TrackOut時間：" + Lots_Info[i].TRACKOUT_TIME + "<br>To5400Time：" + DateTime.Parse(Lots_Info[i].TRACKOUT_TIME).AddSeconds((double)Lots_Info[i].TO_END_STEP_TIME) + "<br>期望交期：" + DateTime.Parse(Lots_Info[i].WIPEND_TIME).ToString() + "','" + ColorTranslator.ToHtml(color) + "', new Date(" + TRACKIN_TIME.Year + "," + (TRACKIN_TIME.Month - 1) + "," + TRACKIN_TIME.Day + "," + TRACKIN_TIME.Hour + "," + TRACKIN_TIME.Minute + "," + TRACKIN_TIME.Second + "," + "), new Date(" + TRACKOUT_TIME.Year + "," + (TRACKOUT_TIME.Month - 1) + "," + TRACKOUT_TIME.Day + "," + TRACKOUT_TIME.Hour + "," + TRACKOUT_TIME.Minute + "," + TRACKOUT_TIME.Second + "," + ")],");
            }
            streamWriter.WriteLine("    ]);");
            streamWriter.WriteLine("var options = {");
            streamWriter.WriteLine("  timeline: {");
            streamWriter.WriteLine("        showRowLabels: true ,");
            streamWriter.WriteLine("    groupByRowLabel: true,");
            streamWriter.WriteLine("    colorByRowLabel: false,");
            streamWriter.WriteLine("    showBarLabels: false,");
            streamWriter.WriteLine("    avoidOverlappingGridLines: false");
            streamWriter.WriteLine("  }");
            streamWriter.WriteLine("    };");
            streamWriter.WriteLine("var observer = new MutationObserver(setBorderRadius);");
            streamWriter.WriteLine("google.visualization.events.addListener(chart, 'ready', function() {");
            streamWriter.WriteLine("    setBorderRadius();");
            streamWriter.WriteLine("    observer.observe(container, {");
            streamWriter.WriteLine("    childList: true,");
            streamWriter.WriteLine("    subtree: true");
            streamWriter.WriteLine("    });");
            streamWriter.WriteLine("});");
            streamWriter.WriteLine("function setBorderRadius()");
            streamWriter.WriteLine("{");
            streamWriter.WriteLine("    Array.prototype.forEach.call(container.getElementsByTagName('rect'), function(rect) {");
            streamWriter.WriteLine("        if (parseFloat(rect.getAttribute('x')) > 0)");
            streamWriter.WriteLine("        {");
            streamWriter.WriteLine("            rect.setAttribute('rx', 5);");
            streamWriter.WriteLine("            rect.setAttribute('ry', 5);");
            streamWriter.WriteLine("        }");
            streamWriter.WriteLine("    });");
            streamWriter.WriteLine("}");
            streamWriter.WriteLine("    dataTable.sort({column: 0, desc: false});");
            streamWriter.WriteLine("chart.draw(dataTable, options);");
            streamWriter.WriteLine("        }");
            streamWriter.WriteLine("</script>");
            streamWriter.WriteLine("<div id = 'GAChart'  style='height: 100%;'></div>");
            streamWriter.WriteLine("<img src='" + FileSaveName + Method + "(Security C).png'/>");
            streamWriter.WriteLine("  <script src = 'https://cdn.jsdelivr.net/npm/bootstrap@5.0.0/dist/js/bootstrap.bundle.min.js'");
            streamWriter.WriteLine("    integrity = 'sha384-p34f1UUtsS3wqzfto5wAAmdvj+osOnFyQFpp4Ua3gs/ZVWx6oOypYoCJhGGScy+8'");
            streamWriter.WriteLine("    crossorigin = 'anonymous' ></script>");
            streamWriter.WriteLine("</body>");
            streamWriter.WriteLine("</html>");
            streamWriter.Close();
            streamWriter.Dispose();
        }
        #endregion

        #region 生成資料到GA專屬資料表
        public async Task<string> InitLotInfoFromCIM(string queryDate, string queryTime, string CATEGORY, string PLANT, string APIUser, string method)
        {
            string sProcID = nameof(InitLotInfoFromCIM);
            _logger.LogProcIn(msMODULE_ID, sProcID);
            string baseUrl = "baseUrl";
            string endpoint = string.Empty;
            if (CATEGORY == "LaserMarking" && method == "FullWIP")
                endpoint = "/GetLMWIP";
            if (CATEGORY == "LaserMarking" && method == "BoatWIP")
                endpoint = "/GetLMWIPBoat";
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseUrl);
                try
                {
                    _logger.Append(msMODULE_ID, sProcID, "準備對" + baseUrl + "呼叫LotInfo api");
                    HttpResponseMessage response = await client.GetAsync($"{endpoint}?QueryDate={queryDate}&QueryTime={queryTime}");
                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResult = await response.Content.ReadAsStringAsync();
                        using JsonDocument doc = JsonDocument.Parse(jsonResult);
                        JsonElement root = doc.RootElement;
                        if (root.TryGetProperty("Success", out JsonElement successElement) && successElement.GetBoolean())
                        {
                            StringBuilder InitLotInfoSQL = new StringBuilder();
                            if (root.TryGetProperty("RawData", out JsonElement rawDataElement) && rawDataElement.ValueKind == JsonValueKind.Array)
                            {
                                if (rawDataElement.GetArrayLength() == 0)
                                {
                                    _logger.Append("Error");
                                    return "Error";
                                }
                                InitLotInfoSQL = new StringBuilder();
                                InitLotInfoSQL.Append(" DELETE SQL ");
                                _logger.Append("準備要下SQL... " + InitLotInfoSQL);
                                DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitLotInfoSQL.ToString());

                                foreach (JsonElement element in rawDataElement.EnumerateArray())
                                {
                                    InitLotInfoSQL = new StringBuilder();
                                    InitLotInfoSQL.Append(" INSERT SQL ");
                                    _logger.Append("準備要下SQL... " + InitLotInfoSQL);
                                    DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitLotInfoSQL.ToString());
                                }
                                _logger.LogProcOut(msMODULE_ID, sProcID);
                                _logger.Append("LotInfo資料生成成功");
                                _logger.LogProcOut(msMODULE_ID, sProcID);
                                return $"LotInfo資料生成成功";
                            }
                            else
                            {
                                _logger.LogProcOut(msMODULE_ID, sProcID);
                                return $"Error: RawData not found or not an array";
                            }
                        }
                        else
                        {
                            _logger.Append(msMODULE_ID, sProcID, $"LotInfo資料錯誤");
                            _logger.LogProcOut(msMODULE_ID, sProcID);
                            return $"Error: {response.Content}：LotInfo資料錯誤";
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogProcOut(msMODULE_ID, sProcID);
                    return $"Exception: {ex.Message}";
                }
            }
            _logger.Append("Lot_Info獲取失敗");
            _logger.LogProcOut(msMODULE_ID, sProcID);
            return $"Lot_Info獲取失敗";
        }
        public async Task<string> InitEQP_STATUSFromCIM(string queryDate, string queryTime, string CATEGORY, string PLANT, string APIUser, string method)
        {
            string sProcID = nameof(InitEQP_STATUSFromCIM);
            _logger.LogProcIn(msMODULE_ID, sProcID);
            string baseUrl = "baseUrl";
            string endpoint = string.Empty;
            if (CATEGORY == "LaserMarking")
                endpoint = "/GetLMEqpStatus";
            using (var client = new HttpClient())
            {
                client.BaseAddress = new Uri(baseUrl);
                try
                {
                    HttpResponseMessage response = await client.GetAsync($"{endpoint}?QueryDate={queryDate}&QueryTime={queryTime}");
                    if (response.IsSuccessStatusCode)
                    {
                        string jsonResult = await response.Content.ReadAsStringAsync();
                        using JsonDocument doc = JsonDocument.Parse(jsonResult);
                        JsonElement root = doc.RootElement;
                        if (root.TryGetProperty("Success", out JsonElement successElement) && successElement.GetBoolean())
                        {
                            StringBuilder InitEQP_STATUSSQL = new StringBuilder();
                            if (root.TryGetProperty("RawData", out JsonElement rawDataElement) && rawDataElement.ValueKind == JsonValueKind.Array)
                            {
                                if (rawDataElement.GetArrayLength() == 0)
                                {
                                    _logger.Append("Error");
                                    return "Error";
                                }
                                InitEQP_STATUSSQL = new StringBuilder();
                                InitEQP_STATUSSQL.Append(" DELETE SQL ");
                                _logger.Append("準備要下SQL... " + InitEQP_STATUSSQL);
                                DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitEQP_STATUSSQL.ToString());

                                foreach (JsonElement element in rawDataElement.EnumerateArray())
                                {
                                    InitEQP_STATUSSQL = new StringBuilder();
                                    InitEQP_STATUSSQL.Append(" INSERT SQL ");
                                    _logger.Append("準備要下SQL... " + InitEQP_STATUSSQL);
                                    DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitEQP_STATUSSQL.ToString());
                                }

                                _logger.Append("EQPInfo資料生成成功");
                                _logger.LogProcOut(msMODULE_ID, sProcID);
                                return $"EQPInfo資料生成成功";
                            }
                            else
                            {
                                return $"Error: RawData not found or not an array";
                            }
                        }
                        else
                        {
                            _logger.Append(msMODULE_ID, sProcID, $"EQP_STATUS資料錯誤");
                            return $"Error: {response.Content}：EQP_STATUS資料錯誤";
                        }
                    }
                }
                catch (Exception ex)
                {
                    return $"Exception: {ex.Message}";
                }
            }
            _logger.Append("EQP_Info獲取失敗");
            _logger.LogProcOut(msMODULE_ID, sProcID);
            return $"EQP_Info獲取失敗";
        }
        public async Task<string> InitMachineInfoFromCIM(string queryDate, string queryTime, string CATEGORY, string PLANT, string APIUser, string method)
        {
            string sProcID = nameof(InitMachineInfoFromCIM);
            _logger.LogProcIn(msMODULE_ID, sProcID);
            string baseUrl = "baseUrl";
            string endpoint = string.Empty;
            string DEVICE_TYPE = string.Empty;
            string MachineID = string.Empty;
            string PRIORITY = string.Empty;
            string Times = string.Empty;
            string Step1WorkTime = string.Empty;
            string Step2WorkTime = string.Empty;
            string Step3WorkTime = string.Empty;
            if (CATEGORY == "LaserMarking")
            {
                Dictionary<(string, string, string), string> DevicePriorityMapping = new Dictionary<(string, string, string), string>();
                endpoint = "/GetLMDeviceMachineMap";
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseUrl);
                    try
                    {
                        HttpResponseMessage response = await client.GetAsync($"{endpoint}?QueryDate={queryDate}&QueryTime={queryTime}");
                        if (response.IsSuccessStatusCode)
                        {
                            string jsonResult = await response.Content.ReadAsStringAsync();
                            using JsonDocument doc = JsonDocument.Parse(jsonResult);
                            JsonElement root = doc.RootElement;
                            if (root.TryGetProperty("Success", out JsonElement successElement) && successElement.GetBoolean())
                            {
                                if (root.TryGetProperty("RawData", out JsonElement rawDataElement) && rawDataElement.ValueKind == JsonValueKind.Array)
                                {
                                    if (rawDataElement.GetArrayLength() == 0)
                                    {
                                        _logger.Append("Error");
                                        return "Error";
                                    }
                                    foreach (JsonElement element in rawDataElement.EnumerateArray())
                                    {
                                        MachineID = element.GetProperty("MachineID").GetString();
                                        DEVICE_TYPE = element.GetProperty("DeviceType").GetString();
                                        PRIORITY = element.GetProperty("Priority").GetDouble().ToString();
                                        Times = element.GetProperty("Times").GetDouble().ToString();
                                        if (Times == "0" && CATEGORY == "LaserMarking") Times = "3900";
                                        if (!DevicePriorityMapping.ContainsKey((MachineID, DEVICE_TYPE, Times)))
                                            DevicePriorityMapping.Add((MachineID, DEVICE_TYPE, Times), PRIORITY);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogProcOut(msMODULE_ID, sProcID);
                        return $"Exception: {ex.Message}";
                    }
                }
                endpoint = "/GetLMCycleTime";
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseUrl);
                    try
                    {
                        HttpResponseMessage response = await client.GetAsync($"{endpoint}?QueryDate={queryDate}&QueryTime={queryTime}");
                        if (response.IsSuccessStatusCode)
                        {
                            string jsonResult = await response.Content.ReadAsStringAsync();
                            using JsonDocument doc = JsonDocument.Parse(jsonResult);
                            JsonElement root = doc.RootElement;
                            if (root.TryGetProperty("Success", out JsonElement successElement) && successElement.GetBoolean())
                            {
                                StringBuilder InitMachineInfoSQL = new StringBuilder();

                                // 處理 RawData
                                if (root.TryGetProperty("RawData", out JsonElement MACHINE_INFOrawDataElement) && MACHINE_INFOrawDataElement.ValueKind == JsonValueKind.Array)
                                {
                                    if (MACHINE_INFOrawDataElement.GetArrayLength() == 0)
                                    {
                                        _logger.Append("Error");
                                        return "Error";
                                    }
                                    InitMachineInfoSQL = new StringBuilder();
                                    InitMachineInfoSQL.Append(" DELETE SQL ");
                                    _logger.Append("準備要下SQL... " + InitMachineInfoSQL);
                                    DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitMachineInfoSQL.ToString());
                                    _logger.Append("結束SQL... " + InitMachineInfoSQL);

                                    foreach (JsonElement element in MACHINE_INFOrawDataElement.EnumerateArray())
                                    {
                                        if (Step1WorkTime != "0")
                                        {
                                            string NowTimes = "-1";
                                            if (Step2WorkTime != "0")
                                            {
                                                NowTimes = "1";
                                            }
                                            if (DevicePriorityMapping.ContainsKey((MachineID, DEVICE_TYPE, "3900")))
                                            {
                                                PRIORITY = DevicePriorityMapping[(MachineID, DEVICE_TYPE, "3900")];
                                            }
                                            else
                                            {
                                                PRIORITY = "1";
                                            }
                                            InitMachineInfoSQL = new StringBuilder();
                                            InitMachineInfoSQL.Append(" INSERT SQL ");
                                            _logger.Append("準備要下SQL... " + InitMachineInfoSQL);
                                            DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitMachineInfoSQL.ToString());

                                        }
                                        if (Step2WorkTime != "0")
                                        {
                                            string NowTimes = "-1";
                                            if (Step3WorkTime != "0")
                                            {
                                                NowTimes = "2";
                                            }
                                            if (DevicePriorityMapping.ContainsKey((MachineID, DEVICE_TYPE, "1")))
                                            {
                                                PRIORITY = DevicePriorityMapping[(MachineID, DEVICE_TYPE, "1")];
                                            }
                                            else
                                            {
                                                PRIORITY = "1";
                                            }
                                            InitMachineInfoSQL = new StringBuilder();
                                            InitMachineInfoSQL.Append(" INSERT SQL ");
                                            _logger.Append("準備要下SQL... " + InitMachineInfoSQL);
                                            DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitMachineInfoSQL.ToString());
                                        }
                                        if (Step3WorkTime != "0")
                                        {
                                            if (DevicePriorityMapping.ContainsKey((MachineID, DEVICE_TYPE, "2")))
                                            {
                                                PRIORITY = DevicePriorityMapping[(MachineID, DEVICE_TYPE, "2")];
                                            }
                                            else
                                            {
                                                PRIORITY = "1";
                                            }
                                            InitMachineInfoSQL = new StringBuilder();
                                            InitMachineInfoSQL.Append(" INSERT SQL ");
                                            _logger.Append("準備要下SQL... " + InitMachineInfoSQL);
                                            DB_Services.Oracle.ExecuteNonQuery(HDB, isProdEnv, InitMachineInfoSQL.ToString());
                                        }
                                    }
                                    _logger.Append("MachineInfo資料生成成功");
                                    _logger.LogProcOut(msMODULE_ID, sProcID);
                                    return $"MachineInfo資料生成成功";
                                }
                                else
                                {
                                    _logger.LogProcOut(msMODULE_ID, sProcID);
                                    return $"Error: RawData not found or not an array";
                                }
                            }
                            else
                            {
                                _logger.Append(msMODULE_ID, sProcID, $"MachineInfo資料錯誤");
                                return $"Error: {response.Content}：MachineInfo資料錯誤";
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogProcOut(msMODULE_ID, sProcID);
                        return $"Exception: {ex.Message}";
                    }
                }


            }
            _logger.Append("CATEGORY參數異常，請檢查");
            _logger.LogProcOut(msMODULE_ID, sProcID);
            return $"CATEGORY參數異常，請檢查";
        }
        #endregion
    }
}
