﻿using System.Security.Principal;

namespace ContactsBook.Models.Domain
{
    public class Contact
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public string Lastname { get; set; }
        public string Email { get; set; }
        public string PhoneNumber { get; set; }
        public string Address { get; set; }


    }
}
