using System;
using System.Reflection.Metadata.Ecma335;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using static System.Net.WebRequestMethods;

public class FilmCatalogue
{
	public string title;
    public int count;
    public List<Film> films;


	public FilmCatalogue(){}

	public FilmCatalogue(string title, int count, List<Film> films)
	{
		this.title = title;
        this.count = count;
        this.films = films;
	}

    //Add film to the end of the list
	public void AddFilm(Film movie)
	{
		films.Add(movie);
		count++;
	}

    //Remove film by index
    public void RemoveFilm(int index)
	{
		if(count > 0)
		{
            films.RemoveAt(index);
            count--;
        }
		else
		{
			Console.WriteLine("Catalogue is empty");
		}
	}

    //Get film by index
	public Film GetFilm(int index)
	{
        if (count > 0)
        {
			return films[index];
        }
        else
        {
            Console.WriteLine("Catalogue is empty");
            return null;
        }
    }

    //Method to get list of films by category
	public List<Film> GetFilmsByDecade(int decade)
	{
		List<Film> newFilms = new List<Film>();
		int thisDecade = 0;
		foreach (Film film in films)
		{
			thisDecade = film.year / 10 * 10;
			if(thisDecade == decade)
			{
				newFilms.Add(film);
			}
		}
		return newFilms;
	}

    public List<Film> GetFilmsByYear(int year)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)
        {
            if (film.year == year)
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

	public List<Film> GetFilmsByLength(int minL, int maxL)
	{
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)
        {
            if (film.length >= minL && film.length <= maxL)
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByRating(int minR, int maxR)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)
         
        {
            if (film.rating >= minR && film.rating <= maxR)
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByLanguage(string lang)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)

        {
            if (film.language.Equals(lang))
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByDirector(string direct)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)

        {
            if (film.director.Equals(direct))
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByGenre(string gen)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)

        {
            if (film.genres.Any(genre => genre.Equals(gen)))
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByActor(string act)
    {
        List<Film> newFilms = new List<Film>();
        foreach (Film film in films)

        {
            if (film.actors.Any(actor => actor.Equals(act)))
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    public List<Film> GetFilmsByLogDate(string date)
    {
        List<Film> newFilms = new List<Film>();
        string[] dateVars = date.Split('-');
        foreach (Film film in films)
        {
            string[] filmVars = film.logDate.Split('-');
            if(dateVars.Length == 1 && filmVars[0].Equals(date))
            {
                newFilms.Add(film);
            }
            if (dateVars.Length == 2 && filmVars[0].Equals(dateVars[0]) && filmVars[1].Equals(dateVars[1]))
            {
                newFilms.Add(film);
            }
            if (dateVars.Length == 3 && filmVars[0].Equals(dateVars[0]) && filmVars[1].Equals(dateVars[1]) && filmVars[2].Equals(dateVars[2]))
            {
                newFilms.Add(film);
            }
        }
        return newFilms;
    }

    //Sorting methods
    public void SortByTitle(string order)
    {
        if(order.Equals("i"))
        {
            films = films.OrderBy(film => film.title).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.title).ToList();
        }
    }

    public void SortByYear(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.year).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.year).ToList();
        }
    }

    public void SortByLength(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.length).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.length).ToList();
        }
    }

    public void SortByRating(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.rating).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.rating).ToList();
        }
    }

    public void SortByLanguage(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.language).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.language).ToList();
        }
    }

    public void SortByDirector(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.director).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.director).ToList();
        }
    }

    public void SortByStatus(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.status).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.status).ToList();
        }
    }

    public void SortByLogDate(string order)
    {
        if (order.Equals("i"))
        {
            films = films.OrderBy(film => film.logDate).ToList();
        }
        else
        {
            films = films.OrderByDescending(film => film.logDate).ToList();
        }
    }

    //Method checks if catalogue contains same movie, if title, year and director are equals
    public bool ContainsFilm(Film filmToCheck)
	{
		foreach(Film film in films)
		{
            if(film.EqualToFilm(filmToCheck))
            {
                return true;
            }
		}
		return false;
	}

	public void PrintToConsole()
	{
        String lineBetween = String.Format("+{0}+{1}+{2}+{3}+{4}+{5}+{6}+{7}+{8}+{9}+{10}+", new String('-', 3), new String('-', 25),
            new String('-', 4), new String('-', 6), new String('-', 6), new String('-', 12), new String('-', 25), new String('-', 45),
            new String('-', 80), new String('-', 13), new String('-', 10));
        String titleLine = String.Format("|{0} |          {1}          |{2,4}|{3,6}|{4,6}|  {5}  |         {6}        |" +
            "                   {7}                    |                                      {8}                                    |" +
            "    {9}   | {10} |", "ID", "Title","Year", "Length", "Rating", "Language", "Director", "Genres", "Actors", "Status", "Log date");
        Console.WriteLine();
        Console.WriteLine("Catalogue title: " + title);
		Console.WriteLine("Number of films in catalogue: " + count);
        Console.WriteLine(); 
        Console.WriteLine(lineBetween);
        Console.WriteLine(titleLine);
        Console.WriteLine(lineBetween);
		for (int i = 0; i < count; i++)
		{
            films[i].id = i + 1;    
			films[i].PrintConsoleLine(lineBetween);
        }
        Console.WriteLine();
    }

	public void PrintToFile()
    {
        ClearFile();
        for (int i = 0; i < count; i++)
        {
            films[i].PrintToFile(title, 3 + i);
		}
    }

    public void ClearFile()
    {
        string path = "./../../../../Catalogues/" + title + ".xlsx";
        FileInfo file = new FileInfo(path);
        ExcelPackage excelFile = new ExcelPackage(file);
        Stream sr = new FileStream(path, FileMode.Open, FileAccess.Read);
        excelFile.Load(sr);
        ExcelWorksheet ws = excelFile.Workbook.Worksheets[0];
        int row = 3;
        while (string.IsNullOrWhiteSpace(ws.Cells[row, 1].Value?.ToString()) == false)
        {
            ws.DeleteRow(row);
        }
        sr.Close();
        excelFile.Save();
    }

    public void ReadFromFile(string path)
    {
        count = 0;
        films = new List<Film>();
        title = path.Substring(25, path.Length - 30);
        FileInfo file = new FileInfo(path);
        ExcelPackage excelFile = new ExcelPackage(file);
        Stream sr = new FileStream(path, FileMode.Open, FileAccess.Read);
        excelFile.Load(sr);
        ExcelWorksheet ws = excelFile.Workbook.Worksheets[0];
        int row = 3;
        while (string.IsNullOrWhiteSpace(ws.Cells[row, 1].Value?.ToString()) == false)
        {
            Film movie = new Film();
            movie.id = int.Parse(ws.Cells[row, 1].Value.ToString());
            movie.title = ws.Cells[row, 2].Value.ToString();
            if (ws.Cells[row, 3].Value.ToString().Equals(""))
            {
                movie.year = -1;
            }
            else
            {
                movie.year = int.Parse(ws.Cells[row, 3].Value.ToString());
            }
            if(ws.Cells[row, 4].Value.ToString().Equals(""))
            {
                movie.length = -1;
            }
            else
            {
                movie.length = int.Parse(ws.Cells[row, 4].Value.ToString());
            }
            if(ws.Cells[row, 5].Value.ToString().Equals(""))
            {
                movie.rating = -1;
            }
            else
            {
                movie.rating = int.Parse(ws.Cells[row, 5].Value.ToString());
            }
            movie.language = ws.Cells[row, 6].Value.ToString();
            movie.director = ws.Cells[row, 7].Value.ToString();
            int i = 0;
            while (string.IsNullOrWhiteSpace(ws.Cells[row, 8 + i].Value?.ToString()) == false && i < 4)
            {
                movie.genres.Add(ws.Cells[row, 8 + i].Value.ToString());
                i++;
            }
            i = 0;
            while (string.IsNullOrWhiteSpace(ws.Cells[row, 12 + i].Value?.ToString()) == false && i < 4)
            {
                movie.actors.Add(ws.Cells[row, 12 + i].Value.ToString());
                i++;
            }
            string line = ws.Cells[row, 16].Value.ToString();
            if(line.Equals("Have seen"))
            {
                movie.status = true;
            }
            else
            {
                movie.status = false;
            }
            movie.logDate = ws.Cells[row, 17].Value.ToString();
            row++;
            count++;
            films.Add(movie);
        }
        sr.Close();
        excelFile.Save();
    }
}
