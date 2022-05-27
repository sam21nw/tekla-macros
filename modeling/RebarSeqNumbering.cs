//
// Purpose:
//		This script assigns sequential numbers (1,2,3...) to reinforcement bars. 
//		The number is stored in a user defined attribute (REBAR_SEQ_NO) and the numbers are
//		unique for similar reinforcement bars in each cast unit mark (=unit having same position)
//
//		To assign the numbers this script will
//			1) Update numbering
//			2) Imports the attributes (no) by using the attribute import


using System;
using System.IO;
using System.Windows.Forms;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using System.Collections.Generic;

namespace Tekla.Technology.Akit.UserScript
{
    public class Script
    {
        public static void Run(Tekla.Technology.Akit.IScript akit)
        {
            akit.Callback("acmd_partnumbers_all", "", "main_frame");
            TSM.Model model = new TSM.Model();

            TSM.ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjects();

            List<Tuple<string, string, int>> rebarDataList = new List<Tuple<string, string, int>>();

            while (moe.MoveNext())
            {
                try
                {
                    TSM.Reinforcement rebar = moe.Current as TSM.Reinforcement;

                    if (rebar != null)
                    {
                        string rebarPos = "", castUnitPos = "";

                        rebar.GetReportProperty("REBAR_POS", ref rebarPos);
                        rebar.GetReportProperty("CAST_UNIT_POS", ref castUnitPos);

                        rebarDataList.Add(new Tuple<string, string, int>(castUnitPos, rebarPos, rebar.Identifier.ID));
                    }
                }
                catch
                {
                }
            }

            rebarDataList.Sort();

            WriteUDA(rebarDataList, model);
        }

        public static void WriteUDA(List<Tuple<string, string, int>> rebarDataList, TSM.Model model)
        {
            int rebarNo = 1;
            string actualRebarNo = "", actualCastUnit = "";

            for (int i = 0; i < rebarDataList.Count; i++)
            {

                if (rebarDataList[i].Item1 != actualCastUnit) //new cast unit
                {
                    actualCastUnit = rebarDataList[i].Item1;
                    actualRebarNo = rebarDataList[i].Item2;
                    rebarNo = 1;
                }
                else if (rebarDataList[i].Item2 != actualRebarNo)//another RebarNo
                {
                    rebarNo++;
                    actualRebarNo = rebarDataList[i].Item2;
                }               

                TSM.ModelObject rebarObject = model.SelectModelObject(new TS.Identifier(rebarDataList[i].Item3));

                rebarObject.SetUserProperty("REBAR_SEQ_NO", rebarNo);
            }
        }
    }
}
