# <a name="install-the-latest-version-of-office-2016"></a>Установка последней версии Office 2016

Подписчики, давшие явное согласие на использование последних сборок Office, первыми получают новые возможности для разработки, включая предварительные версии функций. Чтобы дать согласие на использование последних сборок Office 2016, выполните указанные ниже действия. 

- Если у вас подписка на Office 365 для дома, Office 365 персональный или Office 365 для студентов, прочитайте статью [Примите участие в программе предварительной оценки Office](https://products.office.com/en-us/office-insider).
- Если вы пользуетесь Office 365 для бизнеса, прочитайте статью [Установка сборки раннего выпуска для клиентов Office 365 для бизнеса](https://support.office.com/en-us/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US).
- Если вы используете Office 2016 для Mac:
    - Запустите программу Office 2016 для Mac.
    - Выберите пункт **Проверить наличие обновлений** в меню "Справка".
    - В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office. 

Чтобы получить последнюю сборку, выполните указанные ниже действия. 

1. Скачайте [средство развертывания Office 2016](https://www.microsoft.com/en-us/download/details.aspx?id=49117). 
2. Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.
3. Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml` 

>**Примечание.** На выполнение команды может потребоваться много времени, при этом ход ее выполнения нигде не отображается.

По завершении процесса установки у вас будут последние версии приложений Office 2016. Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл**  >  **Учетная запись**. В разделе "Обновления Office" над номером версии должна быть надпись Office Insiders.

![Снимок экрана, на котором показаны сведения о продукте с надписью "Участники программы предварительной оценки Office"](../../images/officeinsider.PNG)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office

Сведения о минимальных сборках продуктов для каждой платформы см. в следующих статьях:

- [Наборы обязательных элементов API JavaScript для Word](../../reference/requirement-sets/word-api-requirement-sets.md);
- [Наборы обязательных элементов API JavaScript для Excel](../../reference/requirement-sets/excel-api-requirement-sets.md);
- [Наборы обязательных элементов API JavaScript для OneNote](../../reference/requirement-sets/onenote-api-requirement-sets.md);
- [Наборы обязательных элементов Dialog API](../../reference/requirement-sets/dialog-api-requirement-sets.md);
- [Стандартные наборы обязательных элементов API для Office](../../reference/requirement-sets/office-add-in-requirement-sets.md).
