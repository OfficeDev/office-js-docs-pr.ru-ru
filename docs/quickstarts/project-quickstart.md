---
title: Создание первой надстройки области задач Project
description: Узнайте, как создать простую надстройку для области задач Project, используя API JS для Office.
ms.date: 08/04/2021
ms.prod: project
ms.localizationpriority: high
ms.openlocfilehash: cb291a76a97c6cf3c7d816c7c2019337132aecc8
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153950"
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

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-project.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.

- Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.
- Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и клиентским приложением Office.

## <a name="update-the-code"></a>Обновление кода

Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте следующий код в функцию `run`. В этом коде используется API JavaScript для Office, чтобы настроить поле `Name` и поле `Notes` выбранной задачи.

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Запустите локальный веб-сервер.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.

    Выполните следующую команду в корневом каталоге своего проекта. После выполнения этой команды запустится локальный веб-сервер.

    ```command&nbsp;line
    npm run dev-server
    ```

1. В Project создайте простой план проекта.

1. Загрузите свою надстройку в Project, следуя инструкциям в статье [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

1. Выберите отдельную задачу в проекте.

1. В нижней части области задач щелкните ссылку **Выполнить**, чтобы переименовать выбранную задачу и добавить к ней примечания.

    ![Снимок экрана: приложение Project с загруженной надстройкой области задач.](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку области задач Project! Следующим шагом узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.

> [!div class="nextstepaction"]
> [Надстройки Project](../project/project-add-ins.md)

## <a name="see-also"></a>См. также

- [Разработка надстроек Office](../develop/develop-overview.md)
- [Основные принципы надстроек Office](../overview/core-concepts-office-add-ins.md)
