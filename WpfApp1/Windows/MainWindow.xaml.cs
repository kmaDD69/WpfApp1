using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using WorkWithWord.HelperClasses;
using WpfApp1.ModelClasses;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        // Помощник для управления контекстом базы данных
        private ModelEF model;
        // Список всех пользователей
        private List<Users> users;
        // Список всех автомобилей
        private List<Auto> autos;

        // Конструктор главного окна
        public MainWindow()
        {
            InitializeComponent();
            // Создает новый экземпляр контекста базы данных
            model = new ModelEF();
            // Инициализирует списки пользователей и автомобилей пустыми списками
            users = new List<Users>();
            autos = new List<Auto>();
        }

        // Метод для заполнения comboBox'ов данными о пользователях и автомобилях
        private void ComboBoxLoadData()
        {
            // Очищает элементы comboBox'a с пользователями
            comboBoxUsers.Items.Clear();
            // Получает список пользователей данными из базы данных
            users = model.Users.ToList();
            // Добавляет данные о каждом пользователе в comboBox
            foreach (var item in users)
                comboBoxUsers.Items.Add($"{item.FullName} {item.PSeria} {item.PNumber}");
            // Устанавливает первый элемент как выбранный
            comboBoxUsers.SelectedIndex = 0;
            // Получает автомобили текущего выбранного пользователя
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            // Очищает элементы comboBox'a с автомобилями
            comboBoxAutos.Items.Clear();
            // Добавляет данные об автомобилях в comboBox
            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");
            // Устанавливает первый автомобиль как выбранный
            comboBoxAutos.SelectedIndex = 0;
        }

        // Метод вызывается при загрузке окна
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // Загружает данные в comboBox'ы
            ComboBoxLoadData();
        }

        // Метод вызывается при смене выбранного элемента в comboBox'e с пользователями
        private void comboBoxUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Получает автомобили нового выбранного пользователя
            autos = users[comboBoxUsers.SelectedIndex].Auto.ToList();
            // Очищает элементы comboBox'a с автомобилями
            comboBoxAutos.Items.Clear();
            // Добавляет данные об автомобилях в comboBox
            foreach (var item in autos)
                comboBoxAutos.Items.Add($"{item.Model} {item.YearOfRelease.Value.Year} {item.VIN} ");
            // Устанавливает первый автомобиль как выбранный
            comboBoxAutos.SelectedIndex = 0;
        }

        // Метод, обрабатывающий нажатие кнопки 'Сохранить документ'
        private void SaveDocument_Click(object sender, RoutedEventArgs e)
        {
            // Создает диалоговое окно для выбора директории
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            // Выводит описание для окна выбора директории
            fbd.Description = "Выберите место сохранения";
            // Проверяем результаты открытия диалогового окна
            if (System.Windows.Forms.DialogResult.OK == fbd.ShowDialog())
            {
                // Получаем активного пользователя
                Users activeUser = users[comboBoxUsers.SelectedIndex];
                // Получаем активный автомобиль
                Auto activeAuto = activeUser.Auto.ToList()[comboBoxAutos.SelectedIndex];
                // Создаем документ и сохраняем его в указанной директории
                CreateDocument(
                    $"{fbd.SelectedPath}\\{activeUser.FullName}-Автомобиль-{activeUser.FullName}.docx",
                    activeUser,
                    activeAuto);
                // Выводим сообщение о сохранении файла
                System.Windows.MessageBox.Show("Файл сохранён!");
            }
        }

        // Метод создания документа с подстановкой данных пользователя и автомобиля
        private void CreateDocument(string directorypath, Users users, Auto auto)
        {
            // Получаем текущую дату
            var today = DateTime.Now.ToShortDateString();
            // Создаем объект для работы с документом Word
            WordHelper word = new WordHelper("ContractSale.docx");
            // Создаем словарь для замены ключевых слов в документе
            var items = new Dictionary<string, string>
            {
                {"Today", today }, // Замена начального слова <Today> на текущую дату
                {"FullName", users.FullName }, // ФИО
                {"DateOfBirth", users.DateOfBirth.Value.ToShortDateString() }, // Дата рождения
                {"PSeria", users.PSeria.ToString() }, // Серия паспорта
                {"PNumber", users.PNumber.ToString() }, // Номер паспорта
                {"PIDan", users.PVidan }, // Кем выдан паспорт
                // Данные автомобиля
                {"Model", auto.Model }, // Модель автомобиля
                {"Category", auto.Category }, // Категория автомобиля
                //{"Type", auto.Type }, // Тип автомобиля
                {"VIN", auto.VIN }, // VIN номера
                {"RegistrationMark", auto.RegistrationMark }, // Регистрационный знак
                {"YearOfRelease", auto.YearOfRelease.Value.ToString() }, // Год выпуска
                {"EngineNumber", auto.EngineNumber }, // Номер двигателя
                {"Chassis", auto.Chassis }, // Шасси
                {"Bodywork", auto.Bodywork }, // Кузов
                {"Color", auto.Color }, // Цвет
                //{"PTS", auto.PTS }, // ПТС
                //{"NumberPV", auto.NumbePassport }, // Номер ПТС
                //{"VidanPV", auto.VidanPassport } // Кем выдан ПТС
            };
            // Обрабатывает документ, подставляя значения из словаря вместо ключевых слов
            word.Process(items, directorypath);
        }
    }
}
