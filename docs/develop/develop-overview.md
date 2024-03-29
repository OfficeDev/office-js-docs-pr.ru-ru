---
title: Разработка надстроек Office
description: Общие сведения о разработке надстроек Office.
ms.date: 05/25/2022
ms.localizationpriority: high
ms.openlocfilehash: 82573d90f9fa22cb524da01226995e861c258b81
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810024"
---
# <a name="develop-office-add-ins"></a>Разработка надстроек Office

> [!TIP]
> Перед прочтением этой статьи ознакомьтесь с [обзором платформы надстроек Office](../overview/office-add-ins.md).

Все надстройки Office построены на базе платформы надстроек Office. Для каждой создаваемой надстройки следует понять важные принципы, такие как доступность клиентского приложения и платформы, шаблоны программирования API JavaScript для Office, настройку параметров и возможностей надстройки в файле манифеста, способ разработки пользовательского интерфейса и т. д. Эти основные принципы разработки рассматриваются ниже в разделе документации **Жизненный цикл разработки** > **Разработка**. Ознакомьтесь с этими сведениями перед изучением документации для клиентского приложения, надстройку для которого вы создаете (например, [Excel](../excel/index.yml)).

## <a name="create-an-office-add-in"></a>Создание надстройки Office

Можно создать надстройку Office с помощью [генератора Yeoman для надстроек Office](yeoman-generator-overview.md) или Visual Studio.

### <a name="yeoman-generator"></a>Генератор Yeoman

The Yeoman generator for Office Add-ins can be used to create a Node.js Office Add-in project that can be managed with Visual Studio Code or any other editor. The generator can create Office Add-ins for any of the following:

- Excel
- OneNote
- Outlook
- PowerPoint
- Project
- Word
- Пользовательские функции Excel

Создайте проект с помощью HTML, CSS и JavaScript (или TypeScript) либо с помощью Angular или React. Для любой платформы можно также выбирать между JavaScript и Typescript. Дополнительные сведения о создании надстроек с помощью генератора см. в статье [Генератор Yeoman для надстроек Office](yeoman-generator-overview.md).

### <a name="visual-studio"></a>Visual Studio

С помощью Visual Studio можно создавать надстройки Office для Excel, Outlook, Word и PowerPoint. Проект надстройки Office создается в рамках решения Visual Studio и использует HTML, CSS и JavaScript. Дополнительные сведения о создании надстроек с помощью Visual Studio см. в статье [Разработка надстроек Office с помощью Visual Studio](../develop/develop-add-ins-visual-studio.md).

[!include[Yeoman vs Visual Studio comparison](../includes/yeoman-generator-recommendation.md)]

## <a name="understand-the-two-parts-of-an-office-add-in"></a>Общие сведения о двух частях надстройки Office

Надстройка Office состоит из двух частей.

- Манифест надстройки (XML-файл), определяющий параметры и возможности надстройки.

- Веб-приложение, определяющее пользовательский интерфейс и функции компонентов надстройки, таких как области задач, контентные надстройки и диалоговые окна.

The web application uses the Office JavaScript API to interact with content in the Office document where the add-in is running. Your add-in can also do other things that web applications typically do, like call external web services, facilitate user authentication, and more.

### <a name="define-an-add-ins-settings-and-capabilities"></a>Определение параметров и возможностей надстройки

Манифест надстройки Office (XML-файл), определяющий параметры и возможности надстройки. Вы можете настроить манифест, чтобы указать следующие элементы:

- метаданные, описывающие надстройку (например, ИД, версия, описание, отображаемое имя, региональные параметры по умолчанию);
- приложения Office, в которых будет запускаться надстройка;
- разрешения, требующиеся для надстройки;
- способ интеграции надстройки с Office, включая создаваемые ею элементы пользовательского интерфейса (например, настраиваемая вкладка или настраиваемые кнопки на ленте);
- расположение изображений, используемых надстройкой для фирменной символики и значков команд;
- размеры надстройки (например, размеры для контентных надстроек, запрошенная высота для надстроек Outlook);
- правила, определяющие, когда надстройка активируется в контексте сообщения или встречи (только для надстроек Outlook).

Дополнительные сведения о манифесте см. в статье [XML-манифест надстроек Office](add-in-manifests.md).

### <a name="interact-with-content-in-an-office-document"></a>Взаимодействие с содержимым в документе Office

Надстройка Office может использовать API JavaScript для Office, чтобы взаимодействовать с содержимым документа Office, в котором запущена надстройка.

#### <a name="access-the-office-javascript-api-library"></a>Доступ к библиотеке API JavaScript для Office

[!include[information about accessing the Office JS API library](../includes/office-js-access-library.md)]

#### <a name="api-models"></a>Модели API

[!include[information about the Office JS API models](../includes/office-js-api-models.md)]

#### <a name="api-requirement-sets"></a>Наборы обязательных элементов API

[!include[information about the Office JS API requirement sets](../includes/office-js-requirement-sets.md)]

#### <a name="explore-apis-with-script-lab"></a>Изучение API с помощью Script Lab

Script Lab — это надстройка, позволяющая изучать API JavaScript для Office и выполнять фрагменты кода при работе в программах Office, таких как Excel или Word. Она доступна бесплатно в AppSource и является полезным инструментом для добавления в набор средств разработки при создании прототипов и проверке нужных функций в надстройке. В Script Lab можно получить доступ к библиотеке встроенных примеров, чтобы быстро испытать API или использовать пример в качестве отправной точки для собственного кода.

В следующем 1-минутном видео показана надстройка Script Lab в действии.

[![Короткое видео, демонстрирующее работу Script Lab в Excel, Word и PowerPoint.](../images/screenshot-wide-youtube.png 'Ознакомительное видео о Script Lab')](https://aka.ms/scriptlabvideo)

Дополнительные сведения о Script Lab см. в статье [Изучение API JavaScript для Office с помощью Script Lab](../overview/explore-with-script-lab.md).

## <a name="extend-the-office-ui"></a>Настройка интерфейса пользователя Office

Надстройка Office может расширить пользовательский интерфейс Office с помощью команд надстройки и контейнеров HTML, таких как области задач, контентные надстройки и диалоговые окна.

- [Команды надстроек](../design/add-in-commands.md) можно использовать для добавления настраиваемой вкладки, настраиваемых кнопок и меню на стандартную ленту в Office или для расширения стандартного контекстного меню, отображающегося при щелчке правой кнопкой мыши по тексту в документе Office или объекту в Excel. Когда пользователи выбирают команду надстройки, они запускают задачу, определяемую этой командой надстройки, например выполнение кода JavaScript, открытие области задач или запуск диалогового окна.

- Контейнеры HTML, такие как [области задач](../design/task-pane-add-ins.md), [контентные надстройки](../design/content-add-ins.md) и [диалоговые окна](../develop/dialog-api-in-office-add-ins.md) можно использовать для отображения настраиваемого пользовательского интерфейса и предоставления дополнительных функций в приложении Office. Содержимое и функции каждой области задач, контентной надстройки или диалогового окна зависят от указанной вами веб-страницы. Эти веб-страницы могут использовать API JavaScript для Office с целью взаимодействия с содержимым в документе Office, в котором запущена надстройка, а также могут выполнять другие типовые действия веб-страниц, например вызовы внешних веб-служб, упрощение проверки подлинности пользователей и т. д.

На следующем изображении показана команда надстройки на ленте, область задач справа от документа, диалоговое окно или контентная надстройка поверх документа.

![Схема с командами надстроек на ленте, областью задач и диалоговым окном / надстройкой содержимого в документе Office.](../images/add-in-ui-elements.png)

Дополнительные сведения о расширении пользовательского интерфейса Office и разработке интерфейса надстройки см. в статье [Элементы пользовательского интерфейса Office для надстроек Office](../design/interface-elements.md).

## <a name="next-steps"></a>Дальнейшие действия

This article has outlined the different ways to create Office Add-ins, introduced the ways that an add-in can extend the Office UI, described the API sets, and introduced Script Lab as a valuable tool for exploring Office JavaScript APIs and prototyping add-in functionality. Now that you've explored this introductory information, consider continuing your Office Add-ins journey along the following paths.

### <a name="create-an-office-add-in"></a>Создание надстройки Office

Вы можете быстро создать простую надстройку для Excel, OneNote, Outlook, PowerPoint, Project или Word с помощью [5-минутного краткого руководства по началу работы](../index.yml). Если вы уже ознакомились с кратким руководством и хотите создать более сложную надстройку, воспользуйтесь [учебником](../index.yml).

### <a name="learn-more"></a>Дополнительные сведения

Ознакомьтесь с этой документацией, чтобы узнать больше о разработке, тестировании и публикации надстроек Office.

> [!TIP]
> При создании любой надстройки вы можете использовать информацию из раздела [Жизненный цикл разработки](../overview/core-concepts-office-add-ins.md) этой документации, а также сведения из разделов для определенных приложений, соответствующих типу создаваемой надстройки (например, [Excel](../excel/index.yml)).

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Проектирование надстроек Office](../design/add-in-design.md)
- [Тестирование и отладка надстроек Office](../testing/test-debug-office-add-ins.md)
- [Публикация надстроек Office](../publish/publish.md)
