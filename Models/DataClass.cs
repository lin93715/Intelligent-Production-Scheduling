using System;
using System.Collections.Generic;

namespace IPS_API.Models.Data
{
    public class ApiInfoClass
    {
        public bool SUCCESS { get; set; } = false;
        public string ApiName { get; set; }
        public string MachineName { get; set; }
        public string Controller { get; set; }
        public bool Db_IsProd { get; set; }
        public DBChceck dBChceck { get; set; }
        public double ProcessTime { get; set; }
    }

    public class DBChceck
    {
        public DBstatus DB_1 { get; set; }
        public DBstatus DB_2 { get; set; }
    }

    public class DBstatus
    {
        public bool Success { get; set; }
        public string ReturnMsg { get; set; }
    }
    public class APIresponse
    {
        public bool SUCCESS { get; set; } = false;
        public string RETURN_MSG { get; set; } = "";
    }
    public class MBXParam_Response : APIresponse
    {

        public string USER_NAME { get; set; } = "";
        public string MBX_NAME { get; set; } = "";
        public string SERVER_NAME { get; set; } = "";
        public string HOSTNAME { get; set; } = "";
        public string EAP_HOSTNAME { get; set; } = "";

    }


    public class InitialControl_Response : APIresponse
    {
        public string SYSID { get; set; } = "";

    }

    [Serializable]
    public class GAInputData
    {
        public string IPS_ID { get; set; }
        public string CATEGORY { get; set; }
        public float Crossover_Rate { get; set; }
        public float Mutation_Rate { get; set; }
        public int Iteration_Times { get; set; }
        public int Population { get; set; }
        public float Retention_Ratio { get; set; }
        public int AllJobEndTime_Weights { get; set; }
        public int ChangeSetupTime { get; set; }
        public List<CalculationFormula> Calculation_Formula { get; set; }
        public string User { get; set; }
    }
    [Serializable]
    public class CalculationFormula// 適應式計算公式用
    {
        public string Variable_Name { get; set; }//參數名稱
        public int Weights { get; set; }// 權重
    }
    [Serializable]
    public class MachineInfo//機台資訊
    {
        public string CATEGORY { get; set; }
        public string DEVICE_TYPE { get; set; }        
        public string EQP { get; set; }
        public float PREPARE_TIME { get; set; }
        public float TO_END_STEP_TIME { get; set; }
        public int Priority { get; set; }
        public string IPS_STEP { get; set; }
        public string IPS_STEP_NEXT { get; set; }
        public MachineInfo DeepCopy()
        {
            MachineInfo ret = new MachineInfo();
            ret.CATEGORY = CATEGORY;
            ret.DEVICE_TYPE = DEVICE_TYPE;
            ret.EQP = EQP;
            ret.PREPARE_TIME = PREPARE_TIME;
            ret.TO_END_STEP_TIME = TO_END_STEP_TIME;
            ret.Priority = Priority;
            ret.IPS_STEP = IPS_STEP;
            ret.IPS_STEP_NEXT = IPS_STEP_NEXT;
            return ret;
        }
    }
    [Serializable]
    public class LotsInfo
    {
        public string IPS_ID { get; set; }
        public string CATEGORY { get; set; }
        public string LOT_ID { get; set; }              
        public string IPS_STEP { get; set; }            
        public float Qty_pcs { get; set; }              
        public string DEVICE_TYPE { get; set; }         
        public int Priority { get; set; }                         
        public string ISSUE_NUMBER { get; set; }                  
        public string CURRENT_STEP { get; set; }                  
        public float TO_IPS_STEP_TIME { get; set; }               
        public string TRACKIN_TIME { get; set; }                  
        public string EQP { get; set; }                           
        public string TRACKOUT_TIME { get; set; }                 
        public string Memo { get; set; } = "";
        public LotsInfo DeepCopy()
        {
            LotsInfo ret = new LotsInfo();

            ret.IPS_ID = IPS_ID;           
            ret.CATEGORY = CATEGORY;           
            ret.LOT_ID = LOT_ID;
            ret.IPS_STEP = IPS_STEP;
            ret.Qty_pcs = Qty_pcs;
            ret.DEVICE_TYPE = DEVICE_TYPE;
            ret.Priority = Priority;
            ret.ISSUE_NUMBER = ISSUE_NUMBER;
            ret.CURRENT_STEP = CURRENT_STEP;
            ret.TO_IPS_STEP_TIME = TO_IPS_STEP_TIME;
            ret.TRACKIN_TIME = TRACKIN_TIME;
            ret.EQP = EQP;
            ret.TRACKOUT_TIME = TRACKOUT_TIME;
            ret.Memo = Memo;
            return ret;
        }
    }
    public class EQPInfo
    {
        public string CATEGORY { get; set; }
        public string EQP { get; set; }
        public string EQP_STATUS { get; set; }
        public float ETC { get; set; }
        public string IPS_STEP { get; set; }
        public EQPInfo DeepCopy()
        {
            EQPInfo ret = new EQPInfo();
            ret.CATEGORY = CATEGORY;
            ret.EQP = EQP;            
            ret.EQP_STATUS = EQP_STATUS;
            ret.ETC = ETC;
            ret.IPS_STEP = IPS_STEP;          
            return ret;
        }
    }
    public class Chromosome// 染色體(多基因組成)
    {
        public List<float> Genes { get; set; } = new List<float>() { };
        public List<(string, string, string)> ISSUE_NUMBERIndex { get; set; } = new List<(string, string, string)>() { };
        public float AllJobEndTime { get; set; } = 0;
        public float Fitness { get; set; }
        public int ChangeSetupTimeS_COUNT;
        public int EXCEED_WIPEND_TIMES_COUNT;
        public Chromosome()//產生初始染色體
        {
            Genes = new List<float>();
            ISSUE_NUMBERIndex = new List<(string, string, string)>();
            Fitness = 0;
            AllJobEndTime = 0;
            ChangeSetupTimeS_COUNT = 0;
            EXCEED_WIPEND_TIMES_COUNT = 0;
        }
        public Chromosome(int Size)//產生初始染色體(指定基因數量的)
        {
            Genes = new List<float>();
            Random random = new Random((int)DateTime.Now.Ticks);// 隨機物件，用於產生隨機數值
            for (int i = 0; i < Size; i++) Genes.Add((float)random.NextDouble());
            ISSUE_NUMBERIndex = new List<(string, string, string)>();
            Fitness = 0;
            AllJobEndTime = 0;
            ChangeSetupTimeS_COUNT = 0;
            EXCEED_WIPEND_TIMES_COUNT = 0;
        }
        public void SortChromosome(int Size)//將染色體依據Genes大小重新排序，以此排出新的工作順序
        {
            for (int i = 0; i < Size; i++)
            {
                for (int j = i + 1; j < Size; j++)
                {
                    if (Genes[i] > Genes[j])
                    {
                        float TempGene = Genes[i];
                        Genes[i] = Genes[j];
                        Genes[j] = TempGene;
                        (string, string, string) TempLotsIndex = ISSUE_NUMBERIndex[i];
                        ISSUE_NUMBERIndex[i] = ISSUE_NUMBERIndex[j];
                        ISSUE_NUMBERIndex[j] = TempLotsIndex;
                    }
                }
            }
        }
        public Chromosome DeepCopy()
        {
            Chromosome Temp = new Chromosome();
            Temp.Genes = new List<float>();
            foreach (float item in Genes)
                Temp.Genes.Add(item);
            Temp.ISSUE_NUMBERIndex = new List<(string, string, string)>();
            foreach ((string, string, string) item in ISSUE_NUMBERIndex)
                Temp.ISSUE_NUMBERIndex.Add(item);
            Temp.AllJobEndTime = AllJobEndTime;
            Temp.Fitness = Fitness;
            Temp.ChangeSetupTimeS_COUNT = ChangeSetupTimeS_COUNT;
            Temp.EXCEED_WIPEND_TIMES_COUNT = EXCEED_WIPEND_TIMES_COUNT;
            return Temp;
        }
    }
    public class Population// 族群(多染色體組成)
    {
        public List<Chromosome> Chromosomes { get; set; }
        public Population(int Chromosome_Size, int Gene_Size)
        {
            Chromosomes = new List<Chromosome>();
            for (int i = 0; i < Chromosome_Size; i++)// 根據染色體數量產生各條染色體
            {
                Chromosomes.Add(new Chromosome(Gene_Size));
            }
        }
    }
    public class InitData
    {
        public string CATEGORY { get; set; }
        public string QueryDate { get; set; }
        public string QueryTime { get; set; }
        public string APIUser { get; set; }
        public string method { get; set; }
    }
}
