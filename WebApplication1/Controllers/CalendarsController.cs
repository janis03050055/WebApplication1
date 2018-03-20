using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using LinqToExcel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using WebApplication1.Models;

namespace WebApplication1.Views
{
    public class CalendarsController : Controller
    {
        private readonly WebApplication1Context _context;
        private readonly string _folder;
        public CalendarsController(WebApplication1Context context, IHostingEnvironment env)
        {
            _context = context;
            // 把上傳目錄設為：wwwroot\UploadFolder
            _folder = $@"{env.WebRootPath}\UploadFolder";
        }

        // GET: Calendars
        public async Task<IActionResult> Index()
        {
            return View(await _context.Calendar.ToListAsync());
        }

        //回傳Json
        public async Task<IActionResult> JSONdata()
        {
            return Json(await _context.Calendar.ToListAsync());
        }

        // GET: Calendars/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var calendar = await _context.Calendar
                .SingleOrDefaultAsync(m => m.ID == id);
            if (calendar == null)
            {
                return NotFound();
            }

            return View(calendar);
        }

        // GET: Calendars/Create
        public IActionResult Create()
        {
            return View();
        }

        // POST: Calendars/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("ID,date,name,isHoliday,holidayCategory,description")] Calendar calendar)
        {
            if (ModelState.IsValid)
            {
                _context.Add(calendar);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(calendar);
        }

        // GET: Calendars/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var calendar = await _context.Calendar.SingleOrDefaultAsync(m => m.ID == id);
            if (calendar == null)
            {
                return NotFound();
            }
            return View(calendar);
        }

        // POST: Calendars/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("ID,date,name,isHoliday,holidayCategory,description")] Calendar calendar)
        {
            if (id != calendar.ID)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(calendar);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CalendarExists(calendar.ID))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(calendar);
        }

        // GET: Calendars/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var calendar = await _context.Calendar
                .SingleOrDefaultAsync(m => m.ID == id);
            if (calendar == null)
            {
                return NotFound();
            }

            return View(calendar);
        }

        // POST: Calendars/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var calendar = await _context.Calendar.SingleOrDefaultAsync(m => m.ID == id);
            _context.Calendar.Remove(calendar);
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        [HttpPost("UploadFiles")]
        public async Task<IActionResult> Post(List<IFormFile> files)
        {
            //上傳檔案
            if (files.Count < 1)
            {
                return Ok("請選擇檔案，並確定檔案為csv檔。");
            }

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    //儲存檔案
                    var filePath = $@"{_folder}\{file.FileName}";
                    try
                    {
                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            await file.CopyToAsync(stream);
                        }
                    }
                    catch (Exception ex)
                    {
                        return Ok("無法儲存檔案到資料庫，請確定您所輸入的檔案(" + file.FileName + ")，沒有被其他應用程式所開啟或應用。");
                    }
                    
                   
                    //讀取csv檔案
                    var count_line = 0;
                    
                    StreamReader sr = new StreamReader(filePath, Encoding.GetEncoding("Big5"));
                    while (!sr.EndOfStream) // 每次讀取一行，直到檔尾
                    {
                        bool file_success = true;
                        string line = sr.ReadLine();// 讀取文字到 line 變數    
                        string[] dataArray = new string[5];
                        String match_date, match_name, match_isHoliday, match_holidayCategory, match_description;
                        
                        // 判斷第一行是否為正確格式
                        if (count_line == 0 && (line != "\"date\",\"name\",\"isHoliday\",\"holidayCategory\",\"description\"" && line != "date\tname\tisHoliday\tholidayCategory\tdescription"))
                        {
                            return Ok("請確定資料格式是否正確，是否第一列為date、name、isHoliday、holidayCategory、description，並確定檔案為csv檔。");
                        }

                        //有時候csv進行另存時會更改格式，改用tab分割。
                        try
                        {
                            dataArray = line.Split(',');
                            match_date = dataArray[0].Replace("\"", "");
                            match_name = dataArray[1].Replace("\"", "");
                            match_isHoliday = dataArray[2].Replace("\"", "");
                            match_holidayCategory = dataArray[3].Replace("\"", "");
                            match_description = dataArray[4].Replace("\"", "");
                        }
                        catch (Exception ex)
                        {
                            dataArray = line.Split('\t');
                            match_date = dataArray[0].Replace("\"", "");
                            match_name = dataArray[1].Replace("\"", "");
                            match_isHoliday = dataArray[2].Replace("\"", "");
                            match_holidayCategory = dataArray[3].Replace("\"", "");
                            match_description = dataArray[4].Replace("\"", "");
                        }

                        //跳過空白行
                        if (line == ",,,," || line == "\t\t\t\t") file_success = false;

                        if (count_line > 0 && file_success == true)
                        {
                            //確定必填值是否都有資料
                            if (match_date == "") return Ok("請確定輸入的檔案(" + file.FileName + ")第" + count_line + "行的date是不有填寫。");
                            if (match_isHoliday != "是" && match_isHoliday != "否") return Ok("請確定輸入的檔案(" + file.FileName + ")第" + count_line + "行的isHoliday是不有填寫是或否。");
                            if (match_holidayCategory == "") return Ok("請確定輸入的檔案(" + file.FileName + ")第" + count_line + "行的holidayCategory是不有填寫。");
                        }
                        
                        //判斷資料是否格式正確，正確就add一行
                        if (count_line != 0 && file_success == true)
                        {                          
                            try
                            {
                                //將日期正規化
                                DateTime match_date_D = DateTime.ParseExact(dataArray[0].Replace("\"", ""), "yyyy/M/d", System.Globalization.CultureInfo.InvariantCulture);
                                
                                //確定資料庫是否有重複資料，沒有則新增，重複則更新資料
                                if (!CalendarDateExists(match_date_D))
                                {
                                    _context.Calendar.Add(new Calendar()
                                    {
                                        date = match_date_D,
                                        name = match_name,
                                        isHoliday = match_isHoliday,
                                        holidayCategory = match_holidayCategory,
                                        description = match_description
                                    });
                                }
                                else
                                {
                                    _context.Update(new Calendar()
                                    {
                                        date = match_date_D,
                                        name = match_name,
                                        isHoliday = match_isHoliday,
                                        holidayCategory = match_holidayCategory,
                                        description = match_description
                                    });
                                }
                            }
                            catch (Exception ex)
                            {
                                return Ok("請確定輸入的檔案("+ file.FileName + ")第"+ count_line + "行內容的日期格式是否正確須為 yyyy/M/d 。(年為4碼西元年)");
                            }
                            
                        }
                        count_line = count_line + 1;
                    }
                    //全部資料都正確，就savechange
                    await _context.SaveChangesAsync();
                }
            }

            return Ok("資料匯入成功!!");
        }

        private bool CalendarExists(int id)
        {
            return _context.Calendar.Any(e => e.ID == id);
        }

        private bool CalendarDateExists(DateTime date)
        {
            return _context.Calendar.Any(e => e.date == date);
        }
    }
}
