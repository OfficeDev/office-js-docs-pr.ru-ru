# <a name="getstarted-element"></a>Элемент GetStarted

Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).

## <a name="child-elements"></a>Дочерние элементы

| Элемент                       | Обязательный | Description                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [Title](#title)               | Да      | Определяет, где предоставляются функции надстройки.     |
| [Description](#description)   | Да      | URL-адрес файла, содержащего функции JavaScript.|
| [LearnMoreUrl](#learnmoreurl) | Нет       | URL-адрес страницы с подробным описанием надстройки.   |

### <a name="title"></a>Title 

Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).

### <a name="description"></a>Description

Обязательный.  Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).

### <a name="learnmoreurl"></a>LearnMoreUrl

Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый идентификатор элемента **Urls** в разделе [Resources](resources.md).

> [!NOTE]
> В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint. Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен. 

## <a name="see-also"></a>См. также

В следующих примерах кода используется элемент **GetStarted**:

* [Веб-надстройка Excel для работы с форматированием таблиц и диаграмм](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [JavaScript SpecKit для надстроек Word](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
