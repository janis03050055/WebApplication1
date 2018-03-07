using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
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
                    /*
                    //匯入檔案
                    ExcelQueryFactory excel = new ExcelQueryFactory(filePath);
                    excel.AddMapping<Calendar>(e => e.date, "date");
                    excel.AddMapping<Calendar>(e => e.description, "description");
                    excel.AddMapping<Calendar>(e => e.holidayCategory, "holidayCategory");
                    excel.AddMapping<Calendar>(e => e.Id, "Id");
                    excel.AddMapping<Calendar>(e => e.isHoliday, "isHoliday");
                    excel.AddMapping<Calendar>(e => e.name, "name");
                    var exceldata = excel.Worksheet<Calendar>(1);
                    foreach(Calendar itemRow in exceldata)
                    {
                        importCalendar.Append(itemRow.id, itemRow.id, itemRow.id, itemRow.id)
                    }
                    */
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
