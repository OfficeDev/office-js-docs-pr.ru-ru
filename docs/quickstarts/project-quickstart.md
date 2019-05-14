---
title: Создание первой надстройки области задач Project
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: d61f8d83b88dbe69ff0ba9cd4b0afef77a4f03d6
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952254"
---
# <a name="build-your-first-project-task-pane-add-in"></a>Создание первой надстройки области задач Project

В этой статье вы ознакомитесь с процессом создания надстройки для области задач Project.

## <a name="prerequisites"></a>Необходимые компоненты

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Project 2016 или более поздней версии для Windows

## <a name="create-the-add-in"></a>Создание надстройки

1. С помощью генератора Yeoman создайте проект надстройки Project. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

    ```command&nbsp;line
    yo office
    ```

    - **Выберите тип проекта:** `Office Add-in Task Pane project`
    - **Выберите тип сценария:** `Javascript`
    - **Как вы хотите назвать надстройку?** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?** `Project`

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-project.png)
    
    После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.
    
2. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a>Знакомство с проектом

Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач. 

- Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.
- Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.
- Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.
- Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.

## <a name="update-the-code"></a>Обновление кода

Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте приведенный ниже код в пределах функции **run**. В этом коде используется API JavaScript для Office, чтобы настроить поле `Name` и поле `Notes` выбранной задачи.

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

1. Запустите локальный веб-сервер, выполнив следующую команду:

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите `npm start`, примите предложение установить сертификат от генератора Yeoman. 

2. В Project создайте простой план проекта.

3. Загрузите свою надстройку в Project, следуя инструкциям в статье [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

4. Выберите отдельную задачу в проекте.

5. В нижней части области задач щелкните ссылку **Выполнить**, чтобы переименовать выбранную задачу и добавить к ней примечания.

    ![Снимок экрана: приложение Project с загруженной надстройкой области задач](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач Project! Следующим шагом узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.

> [!div class="nextstepaction"]
> [Надстройки Project](../project/project-add-ins.md)

