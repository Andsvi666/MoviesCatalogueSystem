using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using System.ComponentModel;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.DataValidation;

namespace FilmCatalogues
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            List<FilmCatalogue> allLists = ReadAllCatalogues();//reads all current .xlsx files when program starts
            bool running = true;
            while(running)
            {
                PrintMenu();
                ConsoleKeyInfo ans = Console.ReadKey();
                Console.WriteLine();
                int num = 0;
                if(int.TryParse(ans.KeyChar.ToString(), out num))
                {
                    switch(num)
                    {
                        case 1://make film list
                            {
                                allLists = Case1_MakeFilmList(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 2://remove film list
                            {
                                allLists = Case2_RemoveCatalogue(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 3://display catalogue in console
                            {
                                Case3_displayCatalogue(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 4://add new film
                            {
                                allLists = Case4_AddNewFilm(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 5://remove film
                            {
                                allLists = Case5_RemoveFilm(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 6://move film
                            {
                                allLists = Case6_MoveFilm(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 7://Get list by catergory
                            {
                                allLists = Case7_MakeNewCatalogueByCategory(allLists); 
                                ClearConsole(true);
                                break;
                            }
                        case 8://sort catalogue
                            {
                                allLists = Case8_SortCatalogueByCategory(allLists);
                                ClearConsole(true);
                                break;
                            }
                        case 9://close the program
                            {
                                Console.WriteLine("Program closed");
                                running = false;
                                break;
                            }
                        default:
                            {
                                Console.WriteLine("There is no function for this button");
                                ClearConsole(false);
                                break;
                            }
                    }
                }
                else
                {
                    Console.WriteLine("Wrong button, it must be a number button");
                    ClearConsole(false);
                }
            }
        }

        //Case methods
        //Case1
        public static List<FilmCatalogue> Case1_MakeFilmList(List<FilmCatalogue> allLists)
        {
            Console.Write("Write title for the new film catalogue: ");
            string title = Console.ReadLine();
            if(title == null || title.Equals(""))
            {
                Console.WriteLine("You didn't write any title");
            }
            else if (GetAllTitles(allLists).Contains(title))
            {
                Console.WriteLine("There is already catalogue with this title");
            }
            else
            {
                allLists.Add(CreateFilmList(title));
            }
            return allLists;
        }

        //Case2
        public static List<FilmCatalogue> Case2_RemoveCatalogue(List<FilmCatalogue> allLists)
        {
            string request = "From which catalogue do you want to remove: ";
            string title = CheckLists(allLists, request);
            if (title != null)
            {
                if (title == "Watchlist" || title == "Watched Movies")
                {
                    Console.WriteLine("Can't remove standart catalogue - " + title);
                }
                else
                {
                    allLists.Remove(allLists.Find(list => list.title.Equals(title)));
                    //ChangeCatalogueAccessibility(title, true);
                    File.Delete("./../../../../Catalogues/" + title + ".xlsx");
                }
            }
            return allLists;
        }

        //Case3
        public static void Case3_displayCatalogue(List<FilmCatalogue> allLists)
        {
            string request = "Which catalogue do you want to be displayed: ";
            string title = CheckLists(allLists, request);
            if (title != null)
            {
                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
            }
        }

        //Case4
        public static List<FilmCatalogue> Case4_AddNewFilm(List<FilmCatalogue> allLists)
        {
            string request = "To which catalogue do you want to add new film: ";
            string title = CheckLists(allLists, request);
            if (title != null)
            {
                Film newFilm = ReadFilm(title);
                if (!allLists.Find(list => list.title.Equals(title)).ContainsFilm(newFilm))
                {
                    //If you havent seen movie it automaticly adds it to "Watchlist" catalogue
                    if (!newFilm.status && title != "Watchlist")
                    {
                        if (!allLists.Find(list => list.title.Equals("Watchlist")).ContainsFilm(newFilm))
                        {
                            allLists.Find(list => list.title.Equals("Watchlist")).AddFilm(newFilm);
                            newFilm.PrintToFile("Watchlist", 2 + allLists.Find(list => list.title.Equals("Watchlist")).count);
                        }
                    }
                    else if (title != "Watched Movies")//Otherwise it adds to catalogue "WatchedMovies" 
                    {
                        if (!allLists.Find(list => list.title.Equals("Watched Movies")).ContainsFilm(newFilm))
                        {
                            allLists.Find(list => list.title.Equals("Watched Movies")).AddFilm(newFilm);
                            newFilm.PrintToFile("Watched Movies", 2 + allLists.Find(list => list.title.Equals("Watched Movies")).count);
                        }
                    }
                    allLists.Find(list => list.title.Equals(title)).AddFilm(newFilm);
                    newFilm.PrintToFile(title, 2 + allLists.Find(list => list.title.Equals(title)).count);
                }
                else
                {
                    Console.WriteLine("This film is already in the catalogue");
                }
            }
            return allLists;
        }

        //Case5
        public static List<FilmCatalogue> Case5_RemoveFilm(List<FilmCatalogue> allLists)
        {
            string request = "From which catalogue do you want to remove film from: ";
            string title = CheckLists(allLists, request);
            if (title != null)
            {
                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                Console.Write("number of the film you want to remove(1 - " +
                    allLists.Find(list => list.title.Equals(title)).count + "): ");
                string line = Console.ReadLine();
                int id = 0;
                if (Int32.TryParse(line, out id))
                {
                    if (id >= 1 && id <= allLists.Find(list => list.title.Equals(title)).count)
                    {
                        if (title.Equals("Watched Movies") || title.Equals("Watchlist"))
                        {
                            allLists = RemoveFromAllList(allLists, title, id - 1);
                        }
                        else
                        {
                            allLists.Find(list => list.title.Equals(title)).RemoveFilm(id - 1);
                            allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                            allLists.Find(list => list.title.Equals(title)).PrintToFile();
                        }
                    }
                    else
                    {
                        Console.WriteLine("Film number is out of bounds");
                    }
                }
                else
                {
                    Console.WriteLine("Input must be whole number");
                }
            }
            else
            {
                Console.WriteLine("This movie is already in the catalogue");
            }
            return allLists;
        }

        //Case6
        public static List<FilmCatalogue> Case6_MoveFilm(List<FilmCatalogue> allLists)
        {
            string request1 = "From which catalogue do you want to move film: ";
            string title1 = CheckLists(allLists, request1);
            if (title1 != null)
            {
                string request2 = "To which catalogue do you want to move film: ";
                string title2 = CheckLists(allLists, request2);
                if (title2 != null)
                {
                    if (title1.Equals(title2))
                    {
                        Console.WriteLine("You can't move film to the same catalogue");
                    }
                    else if (title1.Equals("Watced Movies") && title2.Equals("Watchlist"))
                    {
                        Console.WriteLine("You can't move film from 'Watched Movies' to 'Watchlist'");
                    }
                    else
                    {
                        allLists.Find(list => list.title.Equals(title1)).PrintToConsole();
                        Console.Write("Number of the film you want to move(1 - " +
                        allLists.Find(list => list.title.Equals(title1)).count + "): ");
                        string line = Console.ReadLine();
                        allLists.Find(list => list.title.Equals(title2)).PrintToConsole();
                        int id = 0;
                        if (Int32.TryParse(line, out id))
                        {
                            if (id >= 1 && id <= allLists.Find(list => list.title.Equals(title1)).count)
                            {
                                Film temp = allLists.Find(list => list.title.Equals(title1)).GetFilm(id - 1);
                                //If movie is moved from  watchlist to catalogue of all seen movies it will change status to seen
                                if (title1 == "Watchlist" && title2 == "Watched Movies")
                                {
                                    allLists = FromWatclistToWatched(temp, allLists, id);
                                }
                                else if (temp != null)
                                {
                                    if (!allLists.Find(list => list.title.Equals(title2)).ContainsFilm(temp))
                                    {
                                        allLists.Find(list => list.title.Equals(title1)).RemoveFilm(id - 1);
                                        allLists.Find(list => list.title.Equals(title1)).PrintToFile();
                                        allLists.Find(list => list.title.Equals(title2)).AddFilm(temp);
                                        allLists.Find(list => list.title.Equals(title2)).PrintToFile();
                                        allLists.Find(list => list.title.Equals(title1)).PrintToConsole();
                                        allLists.Find(list => list.title.Equals(title2)).PrintToConsole();
                                    }
                                    else
                                    {
                                        Console.WriteLine("This film is already in the catalogue");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Film number is out of bounds");
                            }
                        }
                        else
                        {
                            Console.WriteLine("Input must be whole number");
                        }
                    }
                }
            }
            return allLists;
        }

        //Case7
        public static List<FilmCatalogue> Case7_MakeNewCatalogueByCategory(List<FilmCatalogue> allLists)
        {
            Console.Write("Which catagory do you pick: ");
            ConsoleKeyInfo category = Console.ReadKey();
            Console.WriteLine();
            switch (category.KeyChar.ToString())
            {
                case "a"://year
                    {
                        allLists = Case7a_NewCatalogueByYear(allLists);
                        break;
                    }
                case "b"://length
                    {
                        allLists = Case7b_NewCatalogueByLength(allLists);
                        break;
                    }
                case "c"://rating
                    {
                        allLists = Case7c_NewCatalogueByRating(allLists);
                        break;
                    }
                case "d"://language
                    {
                        allLists = Case7d_NewCatalogueByLanguage(allLists);
                        break;
                    }
                case "e"://director
                    {
                        allLists = Case7e_NewCatalogueByDirector(allLists);
                        break;
                    }
                case "f"://genre
                    {
                        allLists = Case7f_NewCatalogueByGenre(allLists);
                        break;
                    }
                case "g"://actor
                    {
                        allLists = Case7g_NewCatalogueByActor(allLists);
                        break;
                    }
                case "h"://Log date
                    {
                        allLists = Case7h_NewCatalogueByLogDate(allLists);
                        break;
                    }
                default:
                    {
                        Console.WriteLine("Wrong button, click one button from one of the given letters");
                        break;
                    }
            }
            return allLists;
        }

        //Case7 methods by category

        //Makes new catalogue by year, case a
        public static List<FilmCatalogue> Case7a_NewCatalogueByYear(List<FilmCatalogue> allLists)
        {
            Console.Write("Do you want to get all movies by decade or year(d/y): ");
            string ch = Console.ReadLine();
            if (ch == "d")
            {
                int currentDecade = DateTime.Now.Year / 10 * 10;
                Console.Write("Which decade(1880 - " + currentDecade + "): ");
                int decade = 0;
                string line = Console.ReadLine();
                if (Int32.TryParse(line, out decade))
                {

                    if (decade < 1880 || decade > currentDecade)
                    {
                        Console.WriteLine("Decade out of bounds");
                    }
                    else if (!line[3].Equals('0'))
                    {
                        Console.WriteLine(line + " is a year not a decade");
                    }
                    else
                    {
                        allLists = Case1_MakeFilmList(allLists);
                        foreach (FilmCatalogue list in allLists)
                        {
                            foreach (Film film in list.GetFilmsByDecade(decade))
                            {
                                if (!allLists.Last().ContainsFilm(film))
                                {
                                    allLists.Last().AddFilm(film);
                                }
                            }
                        }
                        if (allLists.Last().count > 0)
                        {
                            allLists.Last().PrintToConsole();
                            allLists.Last().PrintToFile();
                        }
                        else
                        {
                            File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                            allLists.Remove(allLists.Last());
                            Console.WriteLine("there are no movies from decade " + decade + " in any" +
                                "catalogue. New catalogue removed");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Decade must be a whole number");
                }
            }
            else if (ch == "y")
            {
                int currentYear = DateTime.Now.Year;
                Console.Write("Which year(1888 - " + currentYear + "): ");
                int year = 0;
                string line = Console.ReadLine();
                if (Int32.TryParse(line, out year))
                {
                    if (year < 1888 || year > currentYear)
                    {
                        Console.WriteLine("Year out of bounds");
                    }
                    else
                    {
                        allLists = Case1_MakeFilmList(allLists);
                        foreach (FilmCatalogue list in allLists)
                        {
                            foreach (Film film in list.GetFilmsByYear(year))
                            {
                                if (!allLists.Last().ContainsFilm(film))
                                {
                                    allLists.Last().AddFilm(film);
                                }
                            }
                        }
                        if (allLists.Last().count > 0)
                        {
                            allLists.Last().PrintToConsole();
                            allLists.Last().PrintToFile();
                        }
                        else
                        {
                            File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                            allLists.Remove(allLists.Last());
                            Console.WriteLine("there are no movies from year " + year + " in" +
                                " any catalogue. New catalogue removed");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Decade must be a whole number");
                }
            }
            else
            {
                Console.Write("Input must be 'd' or 'y'");
            }
            return allLists;
        }

        //Makes new catalogue by length, case b
        public static List<FilmCatalogue> Case7b_NewCatalogueByLength(List<FilmCatalogue> allLists)
        {
            Console.Write("Whats the minimum length: ");
            string line1 = Console.ReadLine();
            int minLength = 0;
            if (Int32.TryParse(line1, out minLength))
            {
                Console.Write("Whats the maximum length: ");
                string line2 = Console.ReadLine();
                int maxLength = 0;
                if (Int32.TryParse(line2, out maxLength))
                {
                    if (minLength < 0 || minLength > maxLength)
                    {
                        Console.WriteLine("Length out of bounds");
                    }
                    else
                    {
                        allLists = Case1_MakeFilmList(allLists);
                        foreach (FilmCatalogue list in allLists)
                        {
                            foreach (Film film in list.GetFilmsByLength(minLength, maxLength))
                            {
                                if (!allLists.Last().ContainsFilm(film))
                                {
                                    allLists.Last().AddFilm(film);
                                }
                            }
                        }
                        if (allLists.Last().count > 0)
                        {
                            allLists.Last().PrintToConsole();
                            allLists.Last().PrintToFile();
                        }
                        else
                        {
                            File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                            allLists.Remove(allLists.Last());
                            string temp = minLength + " - " + maxLength;
                            if (minLength == maxLength)
                            {
                                temp = minLength.ToString();
                            }
                            Console.WriteLine("there are no movies that are " + temp + " minutes " +
                                "long in any catalogue. New catalogue removed");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Maximum length must be a whole number");
                }
            }
            else
            {
                Console.WriteLine("Minimum length must be a whole number");
            }
            return allLists;
        }

        //Makes new catalogue by rating, case c
        public static List<FilmCatalogue> Case7c_NewCatalogueByRating(List<FilmCatalogue> allLists)
        {
            Console.Write("Whats the minimum rating: ");
            string line1 = Console.ReadLine();
            int minRating = 0;
            if (Int32.TryParse(line1, out minRating))
            {
                Console.Write("Whats the maximum rating: ");
                string line2 = Console.ReadLine();
                int maxRating = 0;
                if (Int32.TryParse(line2, out maxRating))
                {
                    if (minRating < 0 || minRating > maxRating || maxRating > 100)
                    {
                        Console.WriteLine("Rating out of bounds");
                    }
                    else
                    {
                        allLists = Case1_MakeFilmList(allLists);
                        foreach (FilmCatalogue list in allLists)
                        {
                            foreach (Film film in list.GetFilmsByRating(minRating, maxRating))
                            {
                                if (!allLists.Last().ContainsFilm(film))
                                {
                                    allLists.Last().AddFilm(film);
                                }
                            }
                        }
                        if (allLists.Last().count > 0)
                        {
                            allLists.Last().PrintToConsole();
                            allLists.Last().PrintToFile();
                        }
                        else
                        {
                            File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                            allLists.Remove(allLists.Last());
                            string temp = minRating + " - " + maxRating;
                            if (minRating == maxRating)
                            {
                                temp = minRating.ToString();
                            }
                            Console.WriteLine("there are no movies that have rating of " + temp +
                                " in any catalogue. New catalogue removed");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Maximum length must be a whole number");
                }
            }
            else
            {
                Console.WriteLine("Minimum length must be a whole number");
            }
            return allLists;
        }

        //Makes new catalogue by language, case d
        public static List<FilmCatalogue> Case7d_NewCatalogueByLanguage(List<FilmCatalogue> allLists)
        {
            Console.Write("Which language: ");
            string language = Console.ReadLine();
            if (language == null || language == "")
            {
                Console.WriteLine("Nothing was typed");
            }
            else
            {
                allLists = Case1_MakeFilmList(allLists);
                foreach (FilmCatalogue list in allLists)
                {
                    foreach (Film film in list.GetFilmsByLanguage(language))
                    {
                        if (!allLists.Last().ContainsFilm(film))
                        {
                            allLists.Last().AddFilm(film);
                        }
                    }
                }
                if (allLists.Last().count > 0)
                {
                    allLists.Last().PrintToConsole();
                    allLists.Last().PrintToFile();
                }
                else
                {
                    File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                    allLists.Remove(allLists.Last());
                    Console.WriteLine("there are no movies that have " + language + " language" +
                        " in any catalogue. New catalogue removed");
                }
            }
            return allLists;
        }

        //Makes new catalogue by director, case e
        public static List<FilmCatalogue> Case7e_NewCatalogueByDirector(List<FilmCatalogue> allLists)
        {
            Console.Write("Which director: ");
            string director = Console.ReadLine();
            if (director == null || director == "")
            {
                Console.WriteLine("Nothing was typed");
            }
            else
            {
                allLists = Case1_MakeFilmList(allLists);
                foreach (FilmCatalogue list in allLists)
                {
                    foreach (Film film in list.GetFilmsByDirector(director))
                    {
                        if (!allLists.Last().ContainsFilm(film))
                        {
                            allLists.Last().AddFilm(film);
                        }
                    }
                }
                if (allLists.Last().count > 0)
                {
                    allLists.Last().PrintToConsole();
                    allLists.Last().PrintToFile();
                }
                else
                {
                    File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                    allLists.Remove(allLists.Last());
                    Console.WriteLine("there are no movies directed by " + director +
                        " in any catalogue. New catalogue removed");
                }
            }
            return allLists;
        }

        //Makes new catalogue by genre, case f
        public static List<FilmCatalogue> Case7f_NewCatalogueByGenre(List<FilmCatalogue> allLists)
        {
            Console.Write("Which genre: ");
            string genre = Console.ReadLine();
            if (genre == null || genre == "")
            {
                Console.WriteLine("Nothing was typed");
            }
            else
            {
                allLists = Case1_MakeFilmList(allLists);
                foreach (FilmCatalogue list in allLists)
                {
                    foreach (Film film in list.GetFilmsByGenre(genre))
                    {
                        if (!allLists.Last().ContainsFilm(film))
                        {
                            allLists.Last().AddFilm(film);
                        }
                    }
                }
                if (allLists.Last().count > 0)
                {
                    allLists.Last().PrintToConsole();
                    allLists.Last().PrintToFile();
                }
                else
                {
                    File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                    allLists.Remove(allLists.Last());
                    Console.WriteLine("there are no movies of " + genre + " genre" +
                        " in any catalogue. New catalogue removed");
                }
            }
            return allLists;
        }

        //Makes new catalogue by actor, case g
        public static List<FilmCatalogue> Case7g_NewCatalogueByActor(List<FilmCatalogue> allLists)
        {
            Console.Write("Which actor/actress: ");
            string actor = Console.ReadLine();
            if (actor == null || actor == "")
            {
                Console.WriteLine("Nothing was typed");
            }
            else
            {
                allLists = Case1_MakeFilmList(allLists);
                foreach (FilmCatalogue list in allLists)
                {
                    foreach (Film film in list.GetFilmsByActor(actor))
                    {
                        if (!allLists.Last().ContainsFilm(film))
                        {
                            allLists.Last().AddFilm(film);
                        }
                    }
                }
                if (allLists.Last().count > 0)
                {
                    allLists.Last().PrintToConsole();
                    allLists.Last().PrintToFile();
                }
                else
                {
                    File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                    allLists.Remove(allLists.Last());
                    Console.WriteLine("there are no movies where " + actor + " is performing" +
                        " in any catalogue. New catalogue removed");
                }
            }
            return allLists;
        }

        //Makes new catalogue by log date, case h
        public static List<FilmCatalogue> Case7h_NewCatalogueByLogDate(List<FilmCatalogue> allLists)
        {
            Console.Write("Which year : ");
            string line1 = Console.ReadLine();
            int year = 0;
            if (Int32.TryParse(line1, out year))
            {
                if (year >= 2010 && year <= DateTime.Now.Year)
                {
                    string date = "";
                    Console.Write("Which month(type - to skip): ");
                    string line2 = Console.ReadLine();
                    if (line2.Equals("-"))
                    {
                        date = line1;
                    }
                    else
                    {
                        int month = 0;
                        if (Int32.TryParse(line2, out month))
                        {
                            if (month >= 1 && month <= 12)
                            {
                                Console.Write("Which day(type - to skip): ");
                                string line3 = Console.ReadLine();
                                if (line3.Equals("-"))
                                {
                                    date = line1 + "-" + line2;
                                }
                                else
                                {
                                    int day = 0;
                                    if (Int32.TryParse(line3, out day))
                                    {
                                        if (day >= 1 && day <= DateTime.DaysInMonth(year, month))
                                        {
                                            date = line1 + "-" + line2 + "-" + line3;
                                        }
                                        else
                                        {
                                            Console.WriteLine("day out of bounds");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine("Month out of bounds");
                            }
                        }
                    }
                    allLists = Case1_MakeFilmList(allLists);
                    foreach (FilmCatalogue list in allLists)
                    {
                        foreach (Film film in list.GetFilmsByLogDate(date))
                        {
                            if (!allLists.Last().ContainsFilm(film))
                            {
                                allLists.Last().AddFilm(film);
                            }
                        }
                    }
                    if (allLists.Last().count > 0)
                    {
                        allLists.Last().PrintToConsole();
                        allLists.Last().PrintToFile();
                    }
                    else
                    {
                        File.Delete("./../../../../Catalogues/" + allLists.Last().title + ".txt");
                        allLists.Remove(allLists.Last());
                        Console.WriteLine("there are no movies logged at " + date +
                            " in any catalogue. New catalogue removed");
                    }
                }
                else
                {
                    Console.WriteLine("Year out of bounds");
                }
            }
            else
            {
                Console.WriteLine("Year must be a whole number");
            }
            return allLists;
        }

        //Case8
        public static List<FilmCatalogue> Case8_SortCatalogueByCategory(List<FilmCatalogue> allLists)
        {
            Console.Write("Do you want to sort by increasing or decreasing order(i/d): ");
            string order = Console.ReadLine();
            if (order.Equals("d") || order.Equals("i"))
            {
                string request = "Which catalogue do you want to sort: ";
                string title = CheckLists(allLists, request);
                if (title != null)
                {
                    Console.Write("Which catagory do you pick: ");
                    string category = Console.ReadLine();
                    switch (category)
                    {
                        case "a"://title
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByTitle(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "b"://year
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByYear(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "c"://length
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByLength(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "d"://rating
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByRating(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "e"://language
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByLanguage(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "f"://director
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByDirector(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "g"://status
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByStatus(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        case "h"://log date
                            {
                                allLists.Find(list => list.title.Equals(title)).SortByLogDate(order);
                                allLists.Find(list => list.title.Equals(title)).PrintToConsole();
                                allLists.Find(list => list.title.Equals(title)).PrintToFile();
                                break;
                            }
                        default:
                            {
                                Console.WriteLine("Wrong category letter number");
                                break;
                            }
                    }
                }
            }
            else
            {
                Console.WriteLine("Wrong input");
            }
            return allLists;
        }


        //Other methods
        //Method to print out menu text
        public static void PrintMenu()
        {
            Console.WriteLine("List of functions");
            Console.WriteLine("1. Make new films catalogue");
            Console.WriteLine("2. Remove films catalogue, can't remove 'Watchlist' or 'Watched Movies'");
            Console.WriteLine("3. Display film catalogue");
            Console.WriteLine("4. Add new film to catalogue");
            Console.WriteLine("5. Remove film from catalogue, removing from 'Watchlist' or 'Watched Movies'" +
                " removes from all catalogues");
            Console.WriteLine("6. Move film from one catalogue to another catalogue, moving from 'Watchlist' " +
                "to 'Watched Movies' makes film logged");
            Console.WriteLine("7. Make new catalogue for all films by chosen category: ");
            Console.WriteLine("     a. Year/Decade");
            Console.WriteLine("     b. Length");
            Console.WriteLine("     c. Rating");
            Console.WriteLine("     d. Language");
            Console.WriteLine("     e. Director");
            Console.WriteLine("     f. Genre");
            Console.WriteLine("     g. Actor/Actress");
            Console.WriteLine("     h. Log date");
            Console.WriteLine("8. Sort catalogue by chosen category: ");
            Console.WriteLine("     a. Title");
            Console.WriteLine("     b. Year");
            Console.WriteLine("     c. Length");
            Console.WriteLine("     d. Rating");
            Console.WriteLine("     e. Language");
            Console.WriteLine("     f. Director");
            Console.WriteLine("     g. Status");
            Console.WriteLine("     h. Log date");
            Console.WriteLine("9. Close the program");
            Console.WriteLine(new String('-', 40));
            Console.Write("Which function do you want to initiate: ");
        }

        //Logs movie by moving it from Watchlist to Watched Movies
        public static List<FilmCatalogue> FromWatclistToWatched(Film movie, List<FilmCatalogue> allLists, int id)
        {
            Console.Write("Enter rating from 1 to 100: ");
            string line = Console.ReadLine();
            int rating = 0;
            if (Int32.TryParse(line, out rating))
            {
                if (rating > 0 && rating < 101)
                {
                    foreach (FilmCatalogue list in allLists)
                    {
                        if (list.ContainsFilm(movie))
                        {
                            list.films.Find(film => film.EqualToFilm(movie)).status = true;
                            list.films.Find(film => film.EqualToFilm(movie)).rating = rating;
                            list.films.Find(film => film.EqualToFilm(movie)).logDate =
                                DateTime.Today.Year.ToString() + "-" + DateTime.Today.Month.ToString() + "-" + DateTime.Today.Day.ToString();
                            list.PrintToFile();
                        }
                    }
                    allLists.Find(list => list.title.Equals("Watchlist")).RemoveFilm(id - 1);
                    allLists.Find(list => list.title.Equals("Watchlist")).PrintToFile();
                    allLists.Find(list => list.title.Equals("Watched Movies")).AddFilm(movie);
                    allLists.Find(list => list.title.Equals("Watched Movies")).PrintToFile();
                    allLists.Find(list => list.title.Equals("Watchlist")).PrintToConsole();
                    allLists.Find(list => list.title.Equals("Watched Movies")).PrintToConsole();
                }
                else
                {
                    Console.WriteLine("Rating is out of bounds, must be on scale from 1-100");
                }
            }
            else
            {
                Console.WriteLine("Input must be number");
            }
            return allLists;
        }

        //Remove same film  from  all lists
        public static List<FilmCatalogue> RemoveFromAllList(List<FilmCatalogue> allLists, string title, int id)
        {
            Film movie = allLists.Find(list => list.title.Equals(title)).GetFilm(id);
            foreach (FilmCatalogue list in allLists)
            {
                if(list.ContainsFilm(movie))
                {
                    int newId = list.films.FindIndex(film => film.EqualToFilm(movie)); 
                    list.RemoveFilm(newId);
                    list.PrintToFile();
                }
            }
            return allLists;
        }

        //Method that gets all text file paths and reads all files to list
        public static List<FilmCatalogue> ReadAllCatalogues()
        {
            List<FilmCatalogue> lists = new List<FilmCatalogue>();
            string dir = "./../../../../Catalogues";
            var paths = Directory.EnumerateFiles(dir);
            foreach (string path in paths)
            {
                FilmCatalogue catalogue = new FilmCatalogue();
                catalogue.ReadFromFile(path);
                lists.Add(catalogue);
            }
            return lists;
        }

        //Method gives messege depending on success of program and clears the console by pressing enter
        public static void ClearConsole(bool success)
        {
            if(success)
            {
                Console.Write("Press enter to continue");
            }
            else
            {
                Console.Write("Press enter to try again");
            }
            while(true)
            {
                if(Console.ReadKey(true).Key == ConsoleKey.Enter)
                {
                    Console.Clear();
                    break;
                }
            }
        }

        //Method checks if there are any lists and displays them, then user picks from them
        public static string CheckLists(List<FilmCatalogue> allLists, string request)
        {
            if(allLists.Count == 0)
            {
                Console.WriteLine("There are no catalogues made. First make one");
                return null;
            }
            else
            {
                Console.Write("Available catalogues: ");
                List<string> titles = GetAllTitles(allLists);
                for(int i = 0; i < titles.Count; i++)
                {
                    Console.Write(titles[i]);
                    if(i + 1 < titles.Count)
                    {
                        Console.Write(", ");
                    }
                }
                Console.WriteLine();
                Console.Write(request);
                string ans = Console.ReadLine();
                if(titles.Contains(ans))
                {
                    return ans;
                }
                else
                {
                    Console.WriteLine("There is no catalogue with this title");
                    return null;
                }         
            }
        }

        //Methord return string line of all lists titles
        public static List<string> GetAllTitles(List<FilmCatalogue> allLists)
        {
            List<string> titles = new List<string>();
            foreach (FilmCatalogue list in allLists)
            {
                titles.Add(list.title);
            }
            return titles;
        }

        //Method makes new film catalogue and saves it in excel file
        public static FilmCatalogue CreateFilmList(string title)
        {
            FilmCatalogue list = new FilmCatalogue(title, 0, new List<Film>());
            string path = "./../../../../Catalogues/" + title + ".xlsx";
            FileInfo file = new FileInfo(path);
            ExcelPackage excelFile = new ExcelPackage(file);
            ExcelWorksheet ws = excelFile.Workbook.Worksheets.Add(title);
            ws.Cells["A1"].Value = title;
            ws.Cells["A1:Q1"].Merge = true;
            ws = SetBorders(ws, "A1:Q1", Color.Orange);
            ws = SetStyle(ws, "A1:Q1", Color.RebeccaPurple, Color.Wheat, 40, true);
            ws.Column(1).Width = 6;
            ws.Cells["A2"].Value = "ID";
            ws.Column(2).Width = 25;
            ws.Cells["B2"].Value = "Title";
            ws.Column(3).Width = 10;
            ws.Cells["C2"].Value = "Year";
            ws.Column(4).Width = 15;
            ws.Cells["D2"].Value = "Length";
            ws.Column(5).Width = 13;
            ws.Cells["E2"].Value = "Rating";
            ws.Column(6).Width = 19;
            ws.Cells["F2"].Value = "Language";
            ws.Column(7).Width = 30;
            ws.Cells["G2"].Value = "Director";
            ws.Column(8).Width = 20;
            ws.Column(9).Width = 20;
            ws.Column(10).Width = 20;
            ws.Column(11).Width = 20;
            ws.Cells["H2"].Value = "Genres";
            ws.Cells["H2:K2"].Merge = true;
            ws.Column(12).Width = 30;
            ws.Column(13).Width = 30;
            ws.Column(14).Width = 30;
            ws.Column(15).Width = 30;
            ws.Cells["L2"].Value = "Main Actors";
            ws.Cells["L2:O2"].Merge = true;
            ws.Column(16).Width = 18;
            ws.Cells["P2"].Value = "Status";
            ws.Column(17).Width = 18;
            ws.Cells["Q2"].Value = "Log Date";
            ws = SetBorders(ws, "A2:Q2", Color.Orange);
            ws = SetStyle(ws, "A2:Q2", Color.RebeccaPurple, Color.Wheat, 20, true);
            ws.Columns.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Columns.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            ws.Columns.Style.WrapText = true;
            excelFile.Save();
            Console.WriteLine("new catalogue " + title + " has been made");
            return list;
        }

        //Method sets borders for given cells with given color
        public static ExcelWorksheet SetBorders(ExcelWorksheet ws, string cells, System.Drawing.Color color)
        {
            ws.Cells[cells].Style.Border.Top.Style = ExcelBorderStyle.None;
            ws.Cells[cells].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells[cells].Style.Border.Top.Color.SetColor(color);
            ws.Cells[cells].Style.Border.Bottom.Style = ExcelBorderStyle.None;
            ws.Cells[cells].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells[cells].Style.Border.Bottom.Color.SetColor(color);
            ws.Cells[cells].Style.Border.Left.Style = ExcelBorderStyle.None;
            ws.Cells[cells].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            ws.Cells[cells].Style.Border.Left.Color.SetColor(color);
            ws.Cells[cells].Style.Border.Right.Style = ExcelBorderStyle.None;
            ws.Cells[cells].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells[cells].Style.Border.Right.Color.SetColor(color);
            return ws;
        }

        //Method sets font, font color and background color for given cells
        public static ExcelWorksheet SetStyle(ExcelWorksheet ws, string cells, System.Drawing.Color backgroundColor, System.Drawing.Color fontColor, int fontSize, bool bold)
        {
            ws.Cells[cells].Style.Font.Color.SetColor(fontColor);
            ws.Cells[cells].Style.Font.Size = fontSize;
            ws.Cells[cells].Style.Font.Bold = bold;
            ws.Cells[cells].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[cells].Style.Fill.BackgroundColor.SetColor(backgroundColor);
            return ws;
        }

        //Method makes new Film variable with information read from console
        public static Film ReadFilm(string listType)
        {
            string line;
            Film movie = new Film();
            //Title
            while (true)
            {
                Console.Write("Enter title: ");
                line = Console.ReadLine();
                if (line == "")
                {
                    Console.WriteLine("Nothing was typed");
                }
                else
                {
                    movie.title = line;
                    break;
                }
            }
            //Year
            while (true)
            {
                Console.Write("Enter year(type - to skip): ");
                line = Console.ReadLine();
                int year = 0;
                if(line == "-")
                {
                    movie.year = -1;
                    break;
                }
                if(line == "")
                {
                    Console.WriteLine("Nothing was typed");
                }
                else if(line.Length != 4)
                {
                    Console.WriteLine("Wrong year input format(must be 4 numbers)");
                }
                else if(Int32.TryParse(line, out year))
                {
                    movie.year = year;
                    break;
                }
                else
                {
                    Console.WriteLine("Input must be whole number");
                }
            }
            //Length
            while (true)
            {
                Console.Write("Enter length in minutes(type - to skip): ");
                line = Console.ReadLine();
                int length = 0;
                if (line == "-")
                {
                    movie.length = -1;
                    break;
                }
                if (Int32.TryParse(line, out length))
                {
                    movie.length = length;
                    break;
                }
                else
                {
                    Console.WriteLine("Input must be whole number");
                }
            }
            //Rating
            while (true)
            {
                Console.Write("Enter rating from 1 to 100(type - to skip): ");
                line = Console.ReadLine();
                int rating = 0;
                if (line == "-")
                {
                    movie.rating = -1;
                    break;
                }
                if(Int32.TryParse(line, out rating))
                {
                    if(rating > 0 && rating < 101)
                    {
                        movie.rating = rating;
                        break;
                    }
                    else
                    {
                        Console.WriteLine("Rating is out of bounds, must be on scale from 1-100");
                    }
                }
                else
                {
                    Console.WriteLine("Input must be number");
                }
            }
            //Language
            while(true)
            {
                Console.Write("Enter language(type - to skip): ");
                line = Console.ReadLine();
                if (line == "-")
                {
                    movie.language = "";
                    break;
                }
                if (line == "")
                {
                    Console.WriteLine("Nothing was typed");
                }
                else
                {
                    movie.language = line;
                    break;
                }
            }
            //Director
            while (true)
            {
                Console.Write("Enter director(type - to skip): ");
                line = Console.ReadLine();
                if (line == "-")
                {
                    movie.director = "";
                    break;
                }
                if (line == "")
                {
                    Console.WriteLine("Nothing was typed");
                }
                else
                {
                    movie.director = line;
                    break;
                }
                
            }
            //Genres
            Console.WriteLine("Enter list of genres(maximum 4, write - to end list): ");
            List<String> genres = new List<String>();
            int index = 0;
            while (true)
            {
                index++;
                if (index == 5)
                {
                    Console.WriteLine("Can't add more genres");
                    break;
                }
                line = Console.ReadLine();
                if (line.Equals(""))
                {
                    Console.WriteLine("Nothing was typed");
                    index--;
                }
                else if (line == "-")
                {
                    if(index == 1)
                    {
                        genres.Add("");
                    }
                    break;
                }
                else if(genres.Contains(line))
                {
                    Console.WriteLine("This genre is already added to the film");
                    index--;
                }
                else
                {
                    genres.Add(line);
                }
            }
            movie.genres = genres;
            //Actors/Actresses
            Console.WriteLine("Enter list of main actors/actresses(maximum 4, write - to end list): ");
            List<String> actors = new List<String>();
            index = 0;
            while (true)
            {
                index++;
                if (index == 5)
                {
                    Console.WriteLine("Can't add more actors/actresses");
                    break;
                }
                line = Console.ReadLine();
                if(line.Equals(""))
                {
                    Console.WriteLine("Nothing was typed");
                    index--;
                }
                else if (line == "-")
                {
                    if (index == 1)
                    {
                        actors.Add("");
                    }
                    break;
                }
                else if(actors.Contains(line))
                {
                    Console.WriteLine("This actor/actress is already added to the film");
                    index--;
                }
                else
                {
                    actors.Add(line);
                }
            }
            movie.actors = actors;
            //Status
            if(movie.rating != -1 || listType == "Watched Movies")
            {
                movie.status = true;
            }
            else if (listType == "Watchlist")
            {
                movie.status = false;
            }
            else
            {
                Console.Write("Have you seen this movie?(y/n): ");
                while (true)
                {
                    line = Console.ReadLine();
                    if (line == "y")
                    {
                        movie.status = true;
                        break;
                    }
                    else if (line == "n")
                    {
                        movie.status = false;
                        break;
                    }
                    else if (line == "")
                    {
                        Console.WriteLine("Nothing was typed");
                    }
                    else
                    {
                        Console.WriteLine("Input must be 'y' or 'n'");
                    }
                }
            }
            //Log Date
            if (movie.status == false)
            {
                movie.logDate = "";
            }
            else
            {
                line = DateTime.Today.Year.ToString() + "-" + DateTime.Today.Month.ToString() + "-" + DateTime.Today.Day.ToString();
                movie.logDate = line;
            }
            return movie;
        }
    }
}