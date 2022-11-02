---
title: Обзор платформы надстроек Office
description: Используйте привычные веб-технологии, например HTML, CSS и JavaScript, для взаимодействия с Word, Excel, PowerPoint, OneNote, Project и Outlook, а также для расширения возможностей этих приложений.
ms.date: 04/14/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: 5a780fcc1f863fb6803e2f719fc27338d4a6c366
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810115"
---
# <a name="office-add-ins-platform-overview"></a>Обзор платформы надстроек Office

You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Outlook, Excel, Word, PowerPoint, OneNote, and Project. Your solution can run in Office across multiple platforms, including Windows, Mac, iPad, and in a browser.

![Приложение Office, а также внедренный веб-сайт (надстройка) обеспечивают бесконечные возможности расширения.](../images/addins-overview.png)

Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:

- **Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose functionality from Microsoft and others in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.

- **Создание оснащенных различными функциями интерактивных объектов, которые можно внедрить в документы Office.** Внедряйте карты, диаграммы и интерактивные визуализации, которые пользователи могут добавлять в свои электронные таблицы Excel и презентации PowerPoint.

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a>Чем надстройки Office отличаются от надстроек COM и VSTO?

COM or VSTO add-ins are earlier Office integration solutions that run only in Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the application (for example, Excel), reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.

![Причины использования надстроек Office: кроссплатформенность, централизованное развертывание, простой доступ через AppSource и встроенные стандартные веб-технологии.](../images/why.png)

Преимущества надстроек Office над надстройками, созданными с помощью VBA, модели COM или VSTO.

- Cross-platform support. Office Add-ins run in Office on the web, Windows, Mac, and iPad.

- Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.

- Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.

- Based on standard web technology. You can use any library you like to build Office Add-ins.

## <a name="components-of-an-office-add-in"></a>Компоненты надстройки Office

An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

### <a name="manifest"></a>Манифест

Этот манифест представляет собой XML-файл, который определяет следующие параметры и возможности надстройки:

- Отображаемое имя, описание, идентификатор, версию и языковой стандарт по умолчанию надстройки.

- Способ интеграции надстройки с Office.  

- Уровень разрешений и требования для доступа к данным для надстройки.

### <a name="web-app"></a>Веб-приложение

The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office client application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.

![Компоненты надстройки Hello World.](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a>Расширение возможностей и взаимодействие с клиентами Office

Что позволяют надстройки Office в клиентском приложении Office.

- расширение функциональных возможностей (в любом приложении Office);

- создание новых объектов (Excel или PowerPoint).

### <a name="extend-office-functionality"></a>Расширение функциональных возможностей Office

Добавить новые возможности в приложения Office можно посредством следующего:  

- настраиваемые кнопки ленты и команды меню (в совокупности зовутся "командами надстройки");

- вставляемые области задач.

Пользовательский интерфейс и области задач указаны в манифесте надстройки.  

#### <a name="custom-buttons-and-menu-commands"></a>Настраиваемые кнопки и команды меню  

You can add custom ribbon buttons and menu items to the ribbon in Office on the web and on Windows. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.  

![Настраиваемые кнопки и команды меню.](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a>Области задач  

You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office on iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.

![Использование областей задач в дополнение к командам надстроек.](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a>Расширение возможностей Outlook

Outlook add-ins can extend the Office app ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.

Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification in the Outlook application to provide a seamless experience on the desktop, web, and tablet and mobile devices.

Обзор надстроек Outlook см. в статье [Общие сведения о надстройках Outlook](../outlook/outlook-add-ins-overview.md).

### <a name="create-new-objects-in-office-documents"></a>Создание новых объектов в документах Office

You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.

![Внедрение веб-объектов, называемых контентными надстройками.](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a>API JavaScript для Office

API JavaScript для Office содержат объекты и элементы для создания надстроек и взаимодействия с содержимым Office и веб-службами. Существует общая объектная модель, которую совместно используют Excel, Outlook, Word, PowerPoint, OneNote и Project. Существуют также более обширные объектные модели для Excel и Word, относящиеся к конкретным приложениям. Эти API предоставляют доступ к известным объектам, таким как абзацы и книги, что упрощает создание надстройки для определенного приложения.

## <a name="next-steps"></a>Дальнейшие действия

Дополнительные вводные сведения о разработке надстроек Office см. в статье [Разработка надстроек Office](../develop/develop-overview.md).

## <a name="see-also"></a>См. также

- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
