---
title: Руководство по пользовательским функциям в Excel
description: В этом руководстве вы создадите надстройку Excel, содержащую пользовательскую функцию, которая может выполнять вычисления, запрашивать веб-данные или потоковые веб-данные.
ms.date: 06/10/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 9550986edcbbed56c69e25e183c304ebe6f6cc07
ms.sourcegitcommit: 4f19f645c6c1e85b16014a342e5058989fe9a3d2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/15/2022
ms.locfileid: "66091073"
---
# <a name="tutorial-create-custom-functions-in-excel"></a>Руководство: создание пользовательских функций в Excel

Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`. Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.

В этом руководстве описан порядок выполнения перечисленных ниже задач.
> [!div class="checklist"]
> - Создание надстройки пользовательской функции с помощью [генератора Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).
> - Использование готовой пользовательской функции для выполнения простых вычислений
> - Создание пользовательской функции, которая получает данные из сети Интернет.
> - Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из сети Интернет

## <a name="prerequisites"></a>Предварительные требования

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- Пакет Office, подключенный к подписке Microsoft 365 (включая Office в Интернете).

  > [!NOTE]
  > Если у вас еще нет Office, вы можете [присоединиться к программе для разработчиков Microsoft 365](https://developer.microsoft.com/office/dev-program), чтобы получить бесплатную 90-дневную возобновляемую подписку на Microsoft 365 для использования в процессе разработки.

## <a name="create-a-custom-functions-project"></a>Создание проекта пользовательских функций

 Чтобы начать, создайте проект кода для разработки надстройки пользовательской функции. [Генератор Yeoman для надстроек Office](../develop/yeoman-generator-overview.md) настроит в вашем проекте некоторые готовые пользовательские функции, которые можно попробовать. Если вы уже с помощью краткого руководства по пользовательским функциям создали проект, продолжайте работать с ним и пропустите [этот шаг](#create-a-custom-function-that-requests-data-from-the-web).

> [!NOTE]
> Если повторно создать проект Yo Office, может возникнуть ошибка, так как в кэше Office уже есть экземпляр функции с таким же именем. Это можно предотвратить, [очищая кэш Office](../testing/clear-cache.md) перед запуском `npm run start`.

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Выберите тип проекта:** `Excel Custom Functions Add-in project`
    - **Выберите тип сценария:** `JavaScript`
    - **Как вы хотите назвать надстройку?** `My custom functions add-in`

    :::image type="content" source="../images/yo-office-excel-cf-quickstart.png" alt-text="Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, запрашивающий проекты пользовательских функций.":::

    Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My custom functions add-in"
    ```

1. Выполните построение проекта.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат, предоставленный генератором Yeoman.

1. Запустите локальный веб-сервер, работающий на Node.js. Вы можете попробовать использовать надстройку пользовательской функции в Excel.

# <a name="excel-on-windows-or-mac"></a>[Excel для Windows или Mac](#tab/excel-windows)

Чтобы проверить надстройку в Excel для Windows или Mac, выполните следующую команду. Когда вы выполните эту команду, запустится локальный веб-сервер и откроется приложение Excel, в котором будет загружена ваша надстройка.

```command&nbsp;line
npm run start:desktop
```

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

# <a name="excel-on-the-web"></a>[Excel в Интернете](#tab/excel-online)

Чтобы проверить надстройку в Excel в Интернете, выполните следующую команду. После выполнения этой команды запустится локальный веб-сервер. Замените "{url}" на URL-адрес документа Excel в OneDrive или библиотеке SharePoint, для которой у вас есть разрешения.

[!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

---

## <a name="try-out-a-prebuilt-custom-function"></a>Проверка работы готовой пользовательской функции

Созданный проект пользовательских функций содержит некоторые готовые пользовательские функции, определенные в файле **src/functions/functions.js**. Файл **./manifest.xml** указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`. Вы будете использовать пространство имен CONTOSO для доступа к пользовательским функциям в Excel.

Попробуйте, как работает пользовательская функция `ADD`, выполнив описанные далее шаги.

1. В Excel перейдите в любую ячейку и введите `=CONTOSO`. Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.

1. Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.

Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете и возвращает результат **210**.

[!include[Manually register an add-in](../includes/excel-custom-functions-manually-register.md)]

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a>Создание пользовательской функции, которая запрашивает данные из сети Интернет

Интеграция данных из Интернета — отличный способ расширения функционала Excel через пользовательские функции. Затем вы создадите пользовательскую функцию с именем `getStarCount`, показывающую, сколько звезд имеет данный репозиторий Github.

1. В проекте **Моя надстройка с настраиваемыми функциями** найдите файл **./src/functions/functions.js** и откройте его в редакторе кода.

1. В **function.js** добавьте следующий код.

    ```JS
    /**
      * Gets the star count for a given Github repository.
      * @customfunction 
      * @param {string} userName string name of Github user or organization.
      * @param {string} repoName string name of the Github repository.
      * @return {number} number of stars given to a Github repository.
      */
      async function getStarCount(userName, repoName) {
        try {
          //You can change this URL to any web request you want to work with.
          const url = "https://api.github.com/repos/" + userName + "/" + repoName;
          const response = await fetch(url);
          //Expect that status code is in 200-299 range
          if (!response.ok) {
            throw new Error(response.statusText)
          }
            const jsonResponse = await response.json();
            return jsonResponse.watchers_count;
        }
        catch (error) {
          return error;
        }
      }
    ```

1. Выполните указанную ниже команду, чтобы повторно собрать проект.

    ```command&nbsp;line
    npm run build
    ```

1. Чтобы повторно зарегистрировать надстройку в Excel, выполните указанные ниже действия (для Excel в Интернете, для Windows или для Mac). Выполните описанные ниже действия, чтобы новая функция стала доступной.

### <a name="excel-on-windows-or-mac"></a>[Excel для Windows или Mac](#tab/excel-windows)

1. Закройте Excel, а затем откройте Excel повторно.

1. В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, находящуюся справа от элемента **Мои надстройки**.

    :::image type="content" source="../images/select-insert.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel для Windows с выделенной стрелкой &quot;Мои надстройки&quot;":::

1. В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Моя надстройка с настраиваемыми функциями**, чтобы ее зарегистрировать.

    :::image type="content" source="../images/excel-cf-tutorial-register.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel для Windows с выделенной надстройкой &quot;Пользовательские функции Excel&quot; в списке &quot;Мои надстройки&quot;.":::

# <a name="excel-on-the-web"></a>[Excel в Интернете](#tab/excel-online)

1. В Excel на вкладке **Вставка** выберите пункт **Надстройки**.

    :::image type="content" source="../images/excel-cf-online-register-add-in-1.png" alt-text="Снимок экрана: лента &quot;Вставка&quot; в Excel в Интернете с выделенной кнопкой &quot;Мои надстройки&quot;.":::

1. Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.

1. Выберите **Обзор...** и откройте корневой каталог проекта, созданный генератором Yeoman.

1. Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.

1. Теперь давайте оценим, как работает новая функция. В ячейке **B1** введите текст **=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")** и нажмите клавишу ВВОД. Результат в ячейке **B1** — это текущее количество звезд, отданных репозиторию [Excel-Custom-Functions Github](https://github.com/OfficeDev/Excel-Custom-Functions).

---

## <a name="create-a-streaming-asynchronous-custom-function"></a>Создание потоковой асинхронной пользовательской функции

Функция `getStarCount` возвращает количество звезд, которые есть у репозитория в определенный момент времени. Пользовательские функции также возвращают непрерывно изменяемые данные. Эти функции называются потоковыми передачами функций. Они должны содержать параметр `invocation`, ссылающийся на ячейку, из которой была вызвана функция. Параметр `invocation` используется для обновления содержимого ячейки в любое время.  

В примере кода ниже вы заметите наличие двух функций, `currentTime` и `clock`. Функция `currentTime` — это статическая функция, которая не использует потоковую передачу функций. Она возвращает дату в виде строки. Функция `clock` использует функцию `currentTime` для обеспечения нового времени каждую секунду для ячейки в Excel. В ней используется `invocation.setResult` для передачи времени в ячейку Excel и `invocation.onCanceled` для обработки отмены функции. 

Проект **Моя надстройка с настраиваемыми функциями** уже содержит две следующие функции в файле **./src/functions/functions.js**.

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}
    
/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);
    
  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
```

Чтобы опробовать функции, введите текст **=CONTOSO.CLOCK()** в ячейку **C1** и нажмите ВВОД. Должна отобразиться текущая дата, которая потоком обновляется каждую секунду. Хотя эти часы являются просто таймером в цикле, однако можно использовать аналогичную идею настройки таймера для более сложных функций, которые выполняют веб-запросы в режиме реального времени.

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных. Затем вы можете изменить свой проект, чтобы использовать общую среду выполнения, упрощая взаимодействие с панелью задач. Выполните инструкции из следующей статьи.

> [!div class="nextstepaction"]
> [Настройка надстройки для использования общей среды выполнения](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
