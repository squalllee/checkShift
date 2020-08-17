using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using checkShift.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace checkShift.Factory
{
    class ShiftFactory
    {
        IWorkbook wk = null;
        
        List<ConflictShiftModel> conflictShiftModels = new List<ConflictShiftModel>();

        public ShiftFactory()
        {
            string[] confictshift = { "A01", "C11", "D03" , "D04", "D05", "F07", "F08", "H03", "H04", "H05", "H06", "L05" };
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "E01",
                ConflictShift = new List<string>(confictshift)
            });

            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "E03",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[]{ "H05", "H06"};
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "F04",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "H06" };
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "F05",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "H06", "D03", "F05" };
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "H03",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "A01", "D05", "F04" ,"F05","F07"};
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "H04",
                ConflictShift = new List<string>(confictshift)
            });

            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "H05",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "A01", "D03", "F04", "F05", "F07" ,"F08"};
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "H06",
                ConflictShift = new List<string>(confictshift)
            });

            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "H06",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "A01", "D03", "F04", "F05", "F07", "F08" };
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "L05",
                ConflictShift = new List<string>(confictshift)
            });

            confictshift = new string[] { "A01", "C11", "D05", "F07", "F08", "H03" , "H04", "H05" , "H06", "L05" };
            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "I05",
                ConflictShift = new List<string>(confictshift)
            });

            conflictShiftModels.Add(new ConflictShiftModel
            {
                Shift = "J07",
                ConflictShift = new List<string>(confictshift)
            });


        }
        public List<PersonalShift> ReadShirt(string filePath,DateTime mStartDate, DateTime mEndDate)
        {
            List<PersonalShift> personalShifts = new List<PersonalShift>();

            FileStream fs = File.OpenRead(filePath);

            string extension = Path.GetExtension(filePath);

            int startRowIndex = 4;
            int startColumnIndex = 1;

            DateTime currentMonth = DateTime.Parse(DateTime.Now.Year + "/" + mStartDate.Month.ToString("00") + "/01");
            DateTime endMonth = DateTime.Parse(DateTime.Now.Year + "/" + mEndDate.Month.ToString("00") + "/01");

            int dayInMonth =0;

            if (extension.Equals(".xls"))
            {
                //把xls文件中的數據寫入wk中
                wk = new HSSFWorkbook(fs);
            }
            else
            {
                //把xlsx文件中的數據寫入wk中
                wk = new XSSFWorkbook(fs);
            }

            fs.Close();
            //讀取當前表數據
            ISheet sheet = null;

            IRow row = null;
            
            for(int sheetIndex =0; currentMonth <= endMonth; sheetIndex++)
            {
                sheet = wk.GetSheetAt(sheetIndex);
                dayInMonth = DateTime.DaysInMonth(currentMonth.Year, currentMonth.Month);
                for (int i = startRowIndex; i < sheet.LastRowNum; i++)
                {
                    row = sheet.GetRow(i);
                    if (row == null || row.Cells.Count == 0) continue;
                    for (int j = startColumnIndex; j < startColumnIndex + dayInMonth + 2; j++)
                    {
                        if (j == 1)
                        {
                            if(row.GetCell(startColumnIndex) == null)
                            {
                                break;
                            }
                            if(personalShifts.Where(e => e.UserId == row.GetCell(startColumnIndex).ToString()).Count() == 0)
                            {
                                personalShifts.Add(new PersonalShift
                                {
                                    UserId = row.GetCell(j).ToString(),
                                    UserName = row.GetCell(j + 1).ToString(),
                                    WorkDays = new List<WorkDay>()
                                });
                            }
                           

                            j += 1;
                        }
                        else
                        {
                            PersonalShift personalShift = personalShifts.Where(e => e.UserId == row.GetCell(startColumnIndex).ToString()).Single();
                            if (personalShift != null)
                            {
                                personalShift.WorkDays.Add(new WorkDay
                                {
                                    workDay = DateTime.Parse(currentMonth.ToString("yyyy/MM/") + (j - 2).ToString("00")),
                                    Shift = row.GetCell(j).ToString()
                                });
                            }
                        }
                    }
                }
                currentMonth = currentMonth.AddMonths(1);
                
            }

            return personalShifts;


        }

        public bool checkShift(PersonalShift personalShift, DateTime mStartDate, DateTime mEndDate,bool isCheck8PeriodWork, out string errMsg)
        {
            errMsg = "";
            List<WorkDay> workDays = personalShift.WorkDays.OrderBy(e => e.workDay).ToList();
            int countineWorkDayCount = 0;

            for (int i=0;i<workDays.Count-1;i++) //檢查是否間隔11小時
            {
                ConflictShiftModel conflictShiftModel = conflictShiftModels.Where(e => e.Shift == workDays[i].Shift).FirstOrDefault();
                if (conflictShiftModel != null)
                {
                    if(conflictShiftModel.ConflictShift.Contains(workDays[i+1].Shift))
                    {
                        errMsg = personalShift.UserName + "(" + personalShift.UserId +　") 日期:" + workDays[i].workDay.ToString("yyyy/MM/dd") + "與" + workDays[i+1].workDay.ToString("yyyy/MM/dd") + " 無間隔11小時，請檢查!";
                        return false;
                    }
                }
            }

            foreach(WorkDay workDay in workDays)//檢查是否連七
            {
                if(workDay.Shift.IndexOf("Z") < 0 && workDay.Shift != "")
                {
                    countineWorkDayCount++;
                }
                else
                {
                    countineWorkDayCount = 0;
                }
                
                if(countineWorkDayCount >= 7)
                {
                    errMsg = personalShift.UserName + "(" + personalShift.UserId + ") 於" + workDay.workDay.ToString("yyyy/MM/dd") + " 連續上班七天，請檢查!";
                    return false;
                }  
            }

            if(isCheck8PeriodWork)
            {
                //檢查8週變形工時
                workDays = personalShift.WorkDays.Where(e => e.workDay.CompareTo(DateTime.Parse(mStartDate.ToString("yyyy/MM/dd"))) >= 0 && e.workDay.CompareTo(DateTime.Parse(mEndDate.ToString("yyyy/MM/dd"))) <= 0).OrderBy(e => e.workDay).ToList();

                int HolidayCount = workDays.Where(e => e.Shift.Trim().ToUpper() == "Z01" || e.Shift.Trim().ToUpper() == "Z07").Count();
                if (HolidayCount < 16)
                {
                    errMsg = personalShift.UserName + "(" + personalShift.UserId + ") 違反八週變形工時規定，在週期內休假" + HolidayCount + "天，請檢查!";
                    return false;
                }
            }
            

            return true;
        }
    }
}
