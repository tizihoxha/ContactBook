using ContactsBook.Data;
using ContactsBook.Models;
using ContactsBook.Models.Domain;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.Elfie.Serialization;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Primitives;
using System.Globalization;
using System.Text;
using CsvHelper;
using OfficeOpenXml;

namespace ContactsBook.Controllers
{
    [Authorize]

    public class ContactsController : Controller
    {
        private readonly ContactBookDbContext contactBookDbContext;

        public ContactsController(ContactBookDbContext contactBookDbContext)
        {
            this.contactBookDbContext = contactBookDbContext;
        }

        [HttpGet]

        public async Task<IActionResult> Index()
        {
            var contacts = await contactBookDbContext.Kontaktet.ToListAsync();
            return View(contacts);

        }
        

        [HttpGet("{Id}")]
        public async Task<IActionResult> Index(string contactSearch)
        {
            ViewData["GetContactDetails"] = contactSearch;

            var contacts = from x in contactBookDbContext.Kontaktet select x;

            if (!String.IsNullOrEmpty(contactSearch))
            {
                contacts = contacts.Where(x => x.Name.Contains(contactSearch) || x.Email.Contains(contactSearch));
            }

            var searchResults = await contacts.AsNoTracking().ToListAsync();

            return View(searchResults);
        }

        [HttpGet]
        public async Task<IActionResult> ExportToCsv(string contactSearch)
        {
            var contacts = from x in contactBookDbContext.Kontaktet select x;

            if (!String.IsNullOrEmpty(contactSearch))
            {
                contacts = contacts.Where(x => x.Name.Contains(contactSearch) || x.Email.Contains(contactSearch));
            }

            var searchResults = await contacts.AsNoTracking().ToListAsync();

            var stream = new MemoryStream();
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            using (var csv = new CsvHelper.CsvWriter(writer, CultureInfo.InvariantCulture)) // Use the complete namespace path to CsvWriter
            {
                csv.WriteRecords(searchResults);
                writer.Flush();
                stream.Position = 0;
                var content = await new StreamReader(stream).ReadToEndAsync();
                var bytes = Encoding.UTF8.GetBytes(content);
                return File(bytes, "text/csv", "contacts.csv");
            }
            

         
        }
        // Export search results to Excel file
        [HttpGet]
        public async Task<IActionResult> ExportToExcel(string contactSearch)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var contacts = from x in contactBookDbContext.Kontaktet select x;

            if (!String.IsNullOrEmpty(contactSearch))
            {
                contacts = contacts.Where(x => x.Name.Contains(contactSearch) || x.Email.Contains(contactSearch));
            }

            var searchResults = await contacts.AsNoTracking().ToListAsync();

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Contacts");
                worksheet.Cells.LoadFromCollection(searchResults, true);
                var content = package.GetAsByteArray();
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "contacts.xlsx");
            }
        }

        [HttpGet]
        [Authorize(Roles="Admin")]
        public IActionResult Add()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Add(ContactViewModel addContactRequest)
        {
            var contact = new Contact()
            {
                Id = Guid.NewGuid(),
                Name = addContactRequest.Name,
                Lastname = addContactRequest.Lastname,
                Email = addContactRequest.Email,
                PhoneNumber = addContactRequest.PhoneNumber,
                Address = addContactRequest.Address
            };
            await contactBookDbContext.Kontaktet.AddAsync(contact);
            await contactBookDbContext.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        [HttpGet]
        public async Task<IActionResult> View(Guid id) {
            var contact = await contactBookDbContext.Kontaktet.FirstOrDefaultAsync(x => x.Id == id);
            if (contact != null) 
            {
                var viewmodel = new UpdateViewModel()
                {
                    Id = contact.Id,
                    Name = contact.Name,
                    Lastname = contact.Lastname,
                    Email = contact.Email,
                    PhoneNumber = contact.PhoneNumber,
                    Address = contact.Address

                };
                return await Task.Run(() => View("View", viewmodel));
            }

            return RedirectToAction("Index");

        }

        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> View(UpdateViewModel model)
        {
            var contact = await contactBookDbContext.Kontaktet.FindAsync(model.Id);
            if (contact != null)
            {
                contact.Name = model.Name;
                contact.Lastname = model.Lastname;
                contact.Email = model.Email;
                contact.PhoneNumber = model.PhoneNumber;
                contact.Address = model.Address;

                await contactBookDbContext.SaveChangesAsync();

                return RedirectToAction("Index");
            }
            return RedirectToAction("Index");
        }
        [HttpPost]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> Delete(UpdateViewModel model)
            {
                var contact = await contactBookDbContext.Kontaktet.FindAsync(model.Id);
                if (contact != null)
                {
                    contactBookDbContext.Kontaktet.Remove(contact);
                    await contactBookDbContext.SaveChangesAsync();
                    return RedirectToAction("Index");

                }
                return RedirectToAction("Index");
            }


    }
}
