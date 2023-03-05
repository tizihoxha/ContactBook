using ContactsBook.Models;
using ContactsBook.Models.Domain;
using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;

namespace ContactsBook.Data
{
    public class ContactBookDbContext : IdentityDbContext
    {
        public ContactBookDbContext(DbContextOptions options) : base(options)
        {
        }
        public DbSet<Contact> Kontaktet { get; set; }
        public DbSet<ApplicationUser> ApplicationUser { get;}
    }
}
