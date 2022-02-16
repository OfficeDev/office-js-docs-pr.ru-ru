---
title: Создание первой надстройки области задач OneNote
description: Узнайте, как создать простую надстройку для области задач OneNote, используя API JS для Office.
ms.date: 02/11/2022
ms.prod: onenote
ms.localizationpriority: high
ms.openlocfilehash: 7d806922785f97430619bd74eb04c7c42595aa4e
ms.sourcegitcommit: 61c183a5d8a9d889b6934046c7e4a217dc761b80
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2022
ms.locfileid: "62855578"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a>Создание первой надстройки области задач OneNote

В этой статье вы ознакомитесь с процессом создания надстройки для области задач OneNote.

## <a name="prerequisites"></a>Необходимые компоненты

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project`
- **Выберите тип сценария:** `Javascript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `OneNote`

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-onenote.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.

- Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.
- Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и клиентским приложением Office.

## <a name="update-the-code"></a>Обновление кода

Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте следующий код в функцию `run`. В этом коде используется API JavaScript для OneNote, чтобы настроить заголовок страницы и добавить контур к тексту страницы.

```js
try {
    await OneNote.run(async (context) => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands.
        await context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Запустите локальный веб-сервер. Выполните указанную ниже команду в корневом каталоге своего проекта.

    ```command&nbsp;line
    npm run dev-server
    ```

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Откройте записную книжку в [OneNote в Интернете](https://www.onenote.com/notebooks) и создайте страницу.

1. Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".

    - Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.

    - Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**.

    На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для обычных записных книжек.

    ![Скриншот диалогового окна "Надстройки Office" со вкладкой "Мои надстройки".](../images/onenote-office-add-ins-dialog.png)

1. В диалоговом окне "Отправить надстройку" выберите **manifest.xml** в папке проекта и нажмите кнопку **Отправить**.

1. На вкладке **Главная** ленты нажмите кнопку **Показать область задач**. Область задач надстройки откроется в iFrame рядом со страницей OneNote.

1. В нижней части области задач щелкните ссылку **Выполнить**, чтобы настроить заголовок страницы и добавить контур к тексту страницы.

    ![Снимок экрана с надстройкой, созданной в этом пошаговом руководстве: выделенная кнопка ленты "Показать область задач" и область задач в OneNote.](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач OneNote! Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек OneNote.

> [!div class="nextstepaction"]
> [Обзор API JavaScript для OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Обзор API JavaScript для OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Справочник по API JavaScript для OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
