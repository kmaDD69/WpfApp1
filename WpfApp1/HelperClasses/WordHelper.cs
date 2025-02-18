using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
// WorkWithWord наименование моего проекта
namespace WorkWithWord.HelperClasses
{
    // Класс для работы с документамиWord.
    public class WordHelper
    {
        // Информация о файле.
        private FileInfo _fileinfo;
        // Конструктор класса, принимающий имя файла.
        public WordHelper(string filename)
        {
            // Проверяем существование файла.
            if (File.Exists(filename))
            {
                // Создаём экземпляр FileInfo для данного файла.
                _fileinfo = new FileInfo(filename);
            }
            else
            {
                // Если файл не найден, выбрасывается исключение.
                throw new FileNotFoundException();
            }
        }
        // Метод обработки документа,замены текста и сохранения его под новым именем.
        public void Process(Dictionary<string, string> items, string path)
        {
            // Переменная для хранения приложения Word.
            Word.Application app = null;
            try
            {
                // Создаём новый экземпляр приложения Word.
                app = new Word.Application();
                // Путь к исходному документу.

                object file = _fileinfo.FullName;
                // Значение для необязательных параметров.
                object missing = Type.Missing;
                // Открываем документ.
                app.Documents.Open(file);
                // Проходимся по всем парам ключ-значение.
                foreach (var item in items)
                {
                    // Объект поиска.
                    Word.Find find = app.Selection.Find;
                    // Текст для поиска.
                    find.Text = item.Key;
                    // Текст для замены.
                    find.Replacement.Text = item.Value;
                    // Параметр для продолжения поиска после достижения конца документа.
                    object wrap = Word.WdFindWrap.wdFindContinue;
                    // Параметр для замены всех вхождений.
                    object replace = Word.WdReplace.wdReplaceAll;
                    // Выполняем поиск и замену.
                    find.Execute(
                    FindText: Type.Missing,
                    MatchCase: false,
                    MatchWholeWord: false,
                    MatchWildcards: false,
                    MatchSoundsLike: missing,
                    MatchAllWordForms: false,
                    Forward: true,
                    Wrap: wrap,
                    Format: false,
                    ReplaceWith: missing,
                    Replace: replace);
                }
                // Сохраняем изменённый документ под новым именем.
                app.ActiveDocument.SaveAs2(path);
                // Закрываем открытый документ.
                app.ActiveDocument.Close();
            }
            catch (Exception ex) // Обработка исключений.
            {
                // Выводим сообщение об ошибке.
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // Если приложение было успешно создано.
                if (app != null)
                {
                    // Завершаем работу приложения.
                    app.Quit();
                }
            }
        }
    }
}
