using System;
using OfficeOpenXml;
using System.Drawing;
using System.IO;
using System.Drawing;
using OfficeOpenXml.Style;

public class Film
{
    public int id;
    public string title;
    public int year;
    public int length;//minutes
    public int rating;//1-100
    public string language;
    public string director;
    public List<String> genres = new List<string>();
    public List<String> actors = new List<string>();
    public bool status;//true - seen, false - havent seen
    public string logDate;//Writes date when movie was added 
    public Film() { }
    public Film(int id, string title, int year, int length, int rating, string language,
        string director, List<String> genres, List<String> mainActors, bool status, string logDate)
    {
        this.title = title;
        this.year = year;
        this.length = length;
        this.rating = rating;
        this.language = language;
        this.director = director;
        this.genres = genres;
        this.actors = mainActors;
        this.status = status;
        this.logDate = logDate;
    }

    //Method checks if another given film is equal to this film. Equal if title, year and director are the same
    public bool EqualToFilm(Film anotherFilm)
    {
        if(title.Equals(anotherFilm.title) && year == anotherFilm.year && director.Equals(anotherFilm.director))
        {
            return true;
        }
        return false;
    }

    public void PrintConsoleLine(string lineBetween)
    {

        string genresLine = "";
        string actorsLine = "";
        string statusLine = "";
        for (int i = 0; i < genres.Count; i++)
        {
            genresLine = genresLine + genres[i];
            if(i < genres.Count - 1)
            {
                genresLine = genresLine + ", ";
            }
        }
        for (int i = 0; i < actors.Count; i++)
        {
            actorsLine = actorsLine + actors[i];
            if (i < actors.Count - 1)
            {
                actorsLine = actorsLine + ", ";
            }
        }
        if (status)
        {
            statusLine = "Have seen";
        }
        else
        {
            statusLine = "Have not seen";
        }
        String infoLine = String.Format("|{0,-3}|{1,-25}|{2,-4}|{3,-6}|{4,-6}|{5,-12}|{6,-25}|{7,-45}|{8,-80}|{9,-13}|{10,-10}|", id, title, year, length, rating,
            language, director, genresLine, actorsLine, statusLine, logDate);
        Console.WriteLine(infoLine);
        Console.WriteLine(lineBetween);
    }
        
    //Method prints film to file at given row
    public void PrintToFile(string title, int row)
    {
        id = row - 2;
        string path = "./../../../../Catalogues/" + title + ".xlsx";
        FileInfo file = new FileInfo(path);
        ExcelPackage excelFile = new ExcelPackage(file);
        Stream sr = new FileStream(path, FileMode.Open, FileAccess.Read);
        excelFile.Load(sr);
        ExcelWorksheet ws = excelFile.Workbook.Worksheets[0];
        ws.Cells[row, 1].Value = id;
        ws.Cells[row, 2].Value = this.title;
        if(year == -1)
        {
            ws.Cells[row, 3].Value = "";
        }
        else
        {
            ws.Cells[row, 3].Value = year;
        }
        if (length == -1)
        {
            ws.Cells[row, 4].Value = "";
        }
        else
        {
            ws.Cells[row, 4].Value = length;
        }
        if (rating == -1)
        {
            ws.Cells[row, 5].Value = "";
        }
        else
        {
            ws.Cells[row, 5].Value = rating;
        }
        ws.Cells[row, 6].Value = language;
        ws.Cells[row, 7].Value = director;
        int i = 1;
        foreach(string genre in genres)
        {
            ws.Cells[row, 7 + i].Value = genre;
            i++;
        }
        i = 1;
        foreach(string actor in actors)
        {
            ws.Cells[row, 11 + i].Value = actor;
            i++;
        }
        if(status)
        {
            ws.Cells[row, 16].Value = "Have seen";
        }
        else
        {
            ws.Cells[row, 16].Value = "Have not seen";
        }
        ws.Cells[row, 17].Value = logDate;
        for (int j = 1; j < 18; j++)
        {
            ws.Cells[row, j].Style.Font.Color.SetColor(Color.Wheat);
            ws.Cells[row, j].Style.Font.Size = 13;
            ws.Cells[row, j].Style.Font.Bold = true;
            ws.Cells[row, j].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            if (ws.Cells[row, j].Value == null || ws.Cells[row, j].Value == "")
            {
                ws.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(145, 21, 21));
            }
            else
            {
                ws.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(24, 114, 54));
            }
            ws.Cells[row, j].Style.Border.Top.Style = ExcelBorderStyle.None;
            ws.Cells[row, j].Style.Border.Top.Style = ExcelBorderStyle.Thick;
            ws.Cells[row, j].Style.Border.Top.Color.SetColor(Color.Orange);
            ws.Cells[row, j].Style.Border.Bottom.Style = ExcelBorderStyle.None;
            ws.Cells[row, j].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
            ws.Cells[row, j].Style.Border.Bottom.Color.SetColor(Color.Orange);
            ws.Cells[row, j].Style.Border.Left.Style = ExcelBorderStyle.None;
            ws.Cells[row, j].Style.Border.Left.Style = ExcelBorderStyle.Thick;
            ws.Cells[row, j].Style.Border.Left.Color.SetColor(Color.Orange);
            ws.Cells[row, j].Style.Border.Right.Style = ExcelBorderStyle.None;
            ws.Cells[row, j].Style.Border.Right.Style = ExcelBorderStyle.Thick;
            ws.Cells[row, j].Style.Border.Right.Color.SetColor(Color.Orange);
        }
        sr.Close();
        excelFile.Save();
    }

}
