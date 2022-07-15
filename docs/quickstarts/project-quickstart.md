---
title: Создание первой надстройки области задач Project
description: Узнайте, как создать простую надстройку для области задач Project, используя API JS для Office.
ms.date: 07/13/2022
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: c2f0e31b5a4c958cd155dfeb6d1648f7a2697c69
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797479"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Создание первой надстройки области задач Project

В этой статье вы ознакомитесь с процессом создания надстройки для области задач Project.

## <a name="prerequisites"></a>Необходимые компоненты

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 или более поздней версии для Windows

## <a name="create-the-add-in"></a>Создание надстройки

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project`
- **Выберите тип сценария:** `Javascript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Project`

![Запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-project.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.

- Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.
- Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и клиентским приложением Office. В этом кратком руководстве код `Name` настраивает поле `Notes` и поле выбранной задачи проекта.

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Запустите локальный веб-сервер.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    Выполните следующую команду в корневом каталоге своего проекта. После выполнения этой команды запустится локальный веб-сервер.

    ```command&nbsp;line
    npm run dev-server
    ```

1. В Project создайте простой план проекта.

1. Загрузите свою надстройку в Project, следуя инструкциям в статье [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

1. Выберите отдельную задачу в проекте.

1. В нижней части области задач щелкните ссылку **Выполнить**, чтобы переименовать выбранную задачу и добавить к ней примечания.

    ![Приложение Project с загруженной надстройкой области задач.](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку области задач Project! Следующим шагом узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.

> [!div class="nextstepaction"]
> [Надстройки Project](../project/project-add-ins.md)

## <a name="see-also"></a>См. также

- [Разработка надстроек Office](../develop/develop-overview.md)
- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
- [Использование Visual Studio Code для публикации](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
