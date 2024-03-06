using Acfoss_FYP_Website_Hasif.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Collections.Generic;

namespace Acfoss_FYP_Website_Hasif.Pages
{
    public class PeopleModel : PageModel
    {
        // Change to a static property to simulate data persistence across requests
        public static List<Person> PeopleList { get; set; } = new List<Person>
        {
            new Person { Id = 1, Name = "John", Age = 30 },
            new Person { Id = 2, Name = "Alice", Age = 25 },
            new Person { Id = 3, Name = "Bob", Age = 35 }
        };

        public IActionResult OnGet()
        {
            return Page();
        }

        public IActionResult OnPostDelete(int id)
        {
            // Find the person in the list by Id and remove it
            var personToDelete = PeopleList.Find(p => p.Id == id);
            if (personToDelete != null)
            {
                PeopleList.Remove(personToDelete);
                return RedirectToPage(); // Redirect back to the same page
            }
            else
            {
                return NotFound(); // Return a 404 Not Found if the person to delete is not found
            }
        }

        public IActionResult OnPostEdit(int id, string name, int age)
        {
            // Find the person in the list by Id and update its properties
            var personToEdit = PeopleList.Find(p => p.Id == id);
            if (personToEdit != null)
            {
                personToEdit.Name = name;
                personToEdit.Age = age;
                return RedirectToPage(); // Redirect back to the same page
            }
            else
            {
                return NotFound(); // Return a 404 Not Found if the person to edit is not found
            }
        }
    }
}
