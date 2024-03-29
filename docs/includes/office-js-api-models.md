API JavaScript для Office включает две модели:

- API-интерфейсы для **определенных приложений** предоставляют объекты со строгой типизацией, которые можно использовать для взаимодействия с собственными объектами определенных приложений Office. Например, вы можете использовать API JavaScript для Excel с целью доступа к листам, диапазонам, таблицам, диаграммам и т. д. API для определенных приложений в настоящее время доступны для следующих приложений Office.

    - [Excel](../reference/overview/excel-add-ins-reference-overview.md)
    - [OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
    - [PowerPoint](../reference/overview/powerpoint-add-ins-reference-overview.md)
    - [Word](../reference/overview/word-add-ins-reference-overview.md)

    Эта модель API использует [обещания](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) и позволяет указывать несколько операций в каждом запросе, отправляемом в приложение Office. Подобные пакетные операции могут значительно повысить производительность надстроек для веб-приложений Office. API для определенных приложений появились в Office 2016, и их нельзя использовать для работы с Office 2013.

    > [!NOTE]
    > Существует также специальный API для приложения [Visio](../reference/overview/visio-javascript-reference-overview.md), но его можно использовать только на страницах SharePoint Online для интерактивной работы с встроенными в них диаграммами Visio. Веб-надстройки Office не поддерживаются в Visio.

    Дополнительные сведения об этой модели API см. в статье [Использование модели API для конкретного приложения](../develop/application-specific-api-model.md).

- **Общие** API-интерфейсы можно использовать для доступа к таким компонентам, как пользовательский интерфейс, диалоговые окна и параметры клиентов, общие для нескольких типов приложений Office. Эта модель API использует [обратные вызовы](https://developer.mozilla.org/docs/Glossary/Callback_function), позволяющие указывать только одну операцию в каждом запросе, отправляемом в приложение Office. Общие API появились в Office 2013, и их можно использовать для работы с Office 2013 и более поздними версиями. Подробнее об объектной модели общих API, включающей API для взаимодействия с Outlook, PowerPoint и Project, см. в статье [Объектная модель общих API JavaScript](../develop/office-javascript-api-object-model.md).

> [!NOTE]
>Пользовательские функции без [общей среды выполнения](../testing/runtimes.md#shared-runtime) выполняются только в [среде выполнения JavaScript](../testing/runtimes.md#javascript-only-runtime) , которая определяет приоритет выполнения вычислений. Эти функции используют несколько иную модель программирования.
