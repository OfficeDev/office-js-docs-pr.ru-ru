---
title: Создание автономной надстройки Office на основе кода Script Lab
description: Узнайте, как переместить фрагмент кода из Script Lab в проект Yo Office.
ms.topic: how-to
ms.date: 04/07/2022
ms.localizationpriority: high
ms.openlocfilehash: 725ce9b44c55b46e6d0ab0c085973947fcf88201
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810150"
---
# <a name="create-a-standalone-office-add-in-from-your-script-lab-code"></a>Создание автономной надстройки Office на основе кода Script Lab

Если вы создали фрагмент кода в Script Lab, может потребоваться преобразовать его в автономную надстройку. Вы можете скопировать код из Script Lab в проект, созданный [генератором Yeoman для надстроек Office](../develop/yeoman-generator-overview.md) (другое название — "Yo Office"). Затем вы можете продолжить разработку кода в качестве надстройки, которую в конечном итоге можно развернуть для других пользователей.

Действия в этой статье относятся к [Visual Studio Code](https://code.visualstudio.com/), но вы можете использовать любой другой редактор кода.

## <a name="create-a-new-yo-office-project"></a>Создание проекта Yo Office

Необходимо создать автономный проект надстройки, который будет новым расположением разработки для вашего фрагмента кода.

Запустите команду `yo office --projectType taskpane --ts true --host <host> --name "basic-sample"`, где `<host>` имеет одно из следующих значений.

- excel
- outlook
- powerpoint
- word

> [!IMPORTANT]
> Значение аргумента `--name` должно быть указано в двойных кавычках, даже если оно не содержит пробелов.

Предыдущая команда создает папку проекта с именем **basic-sample**. Он настроен для запуска в указанном вами узле и использует TypeScript. Script Lab использует TypeScript по умолчанию, но большинство фрагментов кода используют JavaScript. При желании вы можете создать проект Yo Office JavaScript. Нужно лишь убедиться, что любой копируемый код является кодом JavaScript.

## <a name="open-the-snippet-in-script-lab"></a>Открытие фрагмента кода в Script Lab

Используйте существующий фрагмент кода в Script Lab, чтобы узнать, как скопировать фрагмент кода в проект, созданный Yo Office.

1. Откройте Office (Word, Excel, PowerPoint или Outlook), а затем — Script Lab.
1. Выберите **Script Lab** > **Код**. Если вы работаете в Outlook, откройте сообщение электронной почты, чтобы увидеть Script Lab на ленте.
1. В области задач Script Lab выберите **Примеры**. Затем выберите базовый пример на основе узла Office, в котором вы работаете.
    - Для Excel или Word выберите пример **Базовый вызов API (TypeScript)**.
    - Для Outlook выберите пример **Использование параметров надстройки**.
    - Для PowerPoint выберите пример **Базовый вызов API (Ofice 2013)**.

## <a name="copy-snippet-code-to-visual-studio-code"></a>Копирование фрагмента кода в Visual Studio Code

Теперь вы можете скопировать код из фрагмента в проект Yo Office в VS Code.

- В VS Code откройте проект **basic-sample**.

На следующих шагах вы скопируете код с нескольких вкладок в Script Lab.

:::image type="content" source="../images/script-lab-script-tabs.png" alt-text="Снимок экрана: вкладки в Script Lab.":::

### <a name="copy-task-pane-code"></a>Копирование кода области задач

1. В VS Code откройте файл **/src/taskpane/taskpane.ts**. Если вы используете проект JavaScript, именем файла будет **taskpane.js**.
1. В Script Lab выберите вкладку **Скрипт**.
1. Скопируйте весь код на вкладке **Скрипт** в буфер обмена. Замените все содержимое **taskpane.ts** (или **taskpane.js** для JavaScript) скопированным кодом.

### <a name="copy-task-pane-html"></a>Копирование HTML-кода области задач

1. В VS Code откройте файл **/src/taskpane/taskpane.html**.
1. В Script Lab выберите вкладку **HTML**.
1. Скопируйте весь HTML-код на вкладке **HTML** в буфер обмена. Замените весь HTML-код внутри тега `<body>` скопированным HTML-кодом.

### <a name="copy-task-pane-css"></a>Копирование CSS области задач

1. В VS Code откройте файл **/src/taskpane/taskpane.css**.
1. В Script Lab выберите вкладку **CSS**.
1. Скопируйте весь код CSS на вкладке **CSS** в буфер обмена. Замените все содержимое **taskpane.css** скопированным кодом CSS.
1. Сохраните все изменения в файлах, обновленных на предыдущих шагах.

## <a name="add-jquery-support"></a>Добавление поддержки jQuery

Script Lab использует jQuery во фрагментах кода. Чтобы успешно выполнить код, необходимо добавить эту зависимость в проект Yo Office.

1. Откройте файл **taskpane.html** и добавьте в раздел `<head>` следующий тег скрипта.

    ```html
     <script src="https://ajax.aspnetcdn.com/ajax/jquery/jquery-3.3.1.js"></script>
    ```

    > [!NOTE]
    > Конкретная версия jQuery может отличаться. Вы можете определить, какую версию использует Script Lab, выбрав вкладку **Библиотеки**.

1. Откройте терминал в VS Code и введите следующие команды.

    ```command&nbsp;line
    npm install --save-dev jquery@3.1.1
    npm install --save-dev @types/jquery@3.3.1
    ```

Если вы создали фрагмент кода с дополнительными зависимостями библиотеки, обязательно добавьте их в проект Yo Office. Найдите список всех зависимостей библиотеки на вкладке **Библиотеки** в Script Lab.

## <a name="handle-initialization"></a>Обработка инициализации

Script Lab автоматически обрабатывает инициализацию `Office.onReady`. Измените код, чтобы предоставить собственный обработчик `Office.onReady`.

1. Откройте файл **taskpane.ts** (или **taskpane.js** для JavaScript).
1. Для Excel или Word замените:

    ```typescript
    $("#run").click(() => tryCatch(run));
    ```

    на:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(() => tryCatch(run));
      });
    });
    ```

1. Для Outlook замените:

    ```typescript
    $("#get").click(get);
    $("#set").click(set);
    $("#save").click(save);
    ```

    на:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#get").click(get);
        $("#set").click(set);
        $("#save").click(save);
      });
    });
    ```

1. Для PowerPoint замените:

    ```typescript
    $("#run").click(run);
    ```

    на:

    ```typescript
    Office.onReady(function () {
      // Office is ready
      $(document).ready(function () {
        // The document is ready
        $("#run").click(run);
      });
    });
    ```

1. Сохраните файл.

## <a name="custom-functions"></a>Настраиваемые функции

Если в вашем фрагменте кода используются настраиваемые функции, необходимо использовать шаблон настраиваемых функций Yo Office. Чтобы превратить настраиваемые функции в автономную надстройку, выполните следующие действия.

1. Выполните команду `yo office --projectType excel-functions --ts true --name "functions-sample"`.

    > [!IMPORTANT]
    > Значение аргумента `--name` должно быть указано в двойных кавычках, даже если оно не содержит пробелов.

1. Откройте Excel, а затем — Script Lab.
1. Выберите **Script Lab** > **Код**.
1. В области задач Script Lab нажмите **Примеры**, а затем выберите пример **Базовая настраиваемая функция**.
1. Откройте файл **/src/functions/functions.ts**. Если вы используете проект JavaScript, именем файла будет **functions.js**.
1. В Script Lab выберите вкладку **Скрипт**.
1. Скопируйте весь код на вкладке **Скрипт** в буфер обмена. Вставьте код в начало **файла functions.ts** (или **functions.js** для JavaScript) с скопированным кодом.
1. Сохраните файл.

## <a name="test-the-standalone-add-in"></a>Тестирование автономной надстройки

После выполнения всех шагов запустите и протестируйте свою автономную надстройку. Выполните следующую команду, чтобы приступить к работе.

```command&nbsp;line
npm start
```

Запустится Office, и вы сможете открыть область задач своей надстройки на ленте. Поздравляем! Теперь вы можете продолжить создание надстройки в качестве автономного проекта.

## <a name="console-logging"></a>Ведение журнала консоли

Многие фрагменты кода в Script Lab записывают выходные данные в раздел консоли в нижней части области задач. В проекте Yo Office нет раздела консоли. Все инструкции `console.log*` будут записываться в консоль отладки по умолчанию (например, в средства разработчика в браузере). Если вы хотите, чтобы выходные данные отправлялись в вашу область задач, обновите код.
