using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.Server.Kestrel.Core.Internal.Http;
using Microsoft.EntityFrameworkCore;
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
        public async Task<IActionResult> Post(List<IFormFile> files, List<Calendar> importCalendar)
        {
            //上傳檔案
            long size = files.Sum(f => f.Length);
            if (files.Count < 1)
            {
                return Content("<script >alert('請選擇檔案');</script >", "text/html");
            }

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    //儲存檔案
                    var filePath = $@"{_folder}\{file.FileName}";
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(stream);
                    }

                    //匯入檔案
                    ExcelQueryFactory excel = new ExcelQueryFactory(filePath);
                    excel.AddMapping<Calendar>(e => e.date, "date");
                    excel.AddMapping<Calendar>(e => e.description, "description");
                    excel.AddMapping<Calendar>(e => e.holidayCategory, "holidayCategory");
                    excel.AddMapping<Calendar>(e => e.ID, "Id");
                    excel.AddMapping<Calendar>(e => e.isHoliday, "isHoliday");
                    excel.AddMapping<Calendar>(e => e.name, "name");

                    var exceldata = excel.Worksheet<Calendar>(1);
                    var importErrorMessages = new List<string>();

                    //檢查資料
                    foreach (var itemRow in exceldata)
                    {
                        var errorMessage = new StringBuilder();
                        var calendarData = new Calendar();

                        calendarData.date = itemRow.date;
                        calendarData.description = itemRow.description;
                        calendarData.name = itemRow.name;

                        //必填不可為空白
                        if (string.IsNullOrWhiteSpace(itemRow.isHoliday))
                        {
                            errorMessage.Append("是否放假 - 不可空白. ");
                        }
                        calendarData.isHoliday = itemRow.isHoliday;

                        if (string.IsNullOrWhiteSpace(itemRow.holidayCategory))
                        {
                            errorMessage.Append("日期類別 - 不可空白. ");
                        }
                        calendarData.holidayCategory = itemRow.holidayCategory;

                        importCalendar.Add(calendarData);

                    }

                }
            }

            return Ok(new { count = files.Count, size });
        }

        private bool CalendarExists(int id)
        {
            return _context.Calendar.Any(e => e.ID == id);
        }

    }
}
