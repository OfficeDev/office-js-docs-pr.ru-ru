---
title: Добавление функций Microsoft Graph в проект быстрого запуска единого входа
description: Узнайте, как добавить новые функции Microsoft Graph к созданной надстройке с поддержкой единого входа.
ms.date: 05/19/2022
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: dbcb32c14824448d2c4309df437c93d01b868288
ms.sourcegitcommit: fcb8d5985ca42537808c6e4ebb3bc2427eabe4d4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2022
ms.locfileid: "65650633"
---
# <a name="add-microsoft-graph-functionality-to-your-sso-quick-start-project"></a>Добавление функций Microsoft Graph в проект быстрого запуска единого входа

> [!IMPORTANT]
> Эта статья создана на основе надстройки с поддержкой единого входа, созданной с помощью краткого руководства по [единому входу](sso-quickstart.md). Прежде чем прочитать эту статью, выполните инструкции из краткого руководства.

В [кратком](sso-quickstart.md) руководстве по единому входу создается надстройка с поддержкой единого входа, которая получает сведения о профиле вошедвшего пользователя и записывает их в документ или сообщение. В этой статье описан процесс обновления надстройки, созданной с помощью генератора Yeoman в кратком руководстве по единому входу, чтобы добавить новые функциональные возможности, которые требуют различных разрешений.

## <a name="prerequisites"></a>Предварительные требования

- Надстройка Office, которую вы создали, следуя инструкциям в кратком руководстве по [единому входу](sso-quickstart.md).

- По крайней мере несколько файлов и папок, OneDrive для бизнеса в вашей Microsoft 365 подписке.

- [Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases)).

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Просмотр содержимого проекта

Начнем с краткой проверки проекта надстройки, созданного ранее с [помощью генератора Yeoman](sso-quickstart.md).

> [!NOTE]
> В местах, где эта статья ссылается на **** файлы скриптов с.jsфайла, предположим, что расширение **TS-файла** используется, если проект был создан с помощью TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Добавление новых функций

Надстройка, созданная с помощью краткого руководства по единому входу, использует Microsoft Graph для получения сведений о профиле вошедщего пользователя и записи этой информации в документ или сообщение. Давайте изменяем функциональные возможности надстройки таким образом, чтобы она получает имена первых 10 файлов и папок из учетной записи пользователя, выполнившего вход OneDrive для бизнеса записывает эти сведения в документ или сообщение. Для включения этой новой функции требуется обновление разрешений приложения в Azure и обновление кода в проекте надстройки.

### <a name="update-app-permissions-in-azure"></a>Обновление разрешений приложения в Azure

Прежде чем надстройка сможет успешно прочитать содержимое учетной записи OneDrive для бизнеса, сведения о регистрации приложения в Azure должны быть обновлены соответствующими разрешениями. Выполните следующие действия, чтобы предоставить приложению разрешение **Files.Read.All** и отозвать разрешение **User.Read** , которое больше не требуется.

1. Войдите [в портал Azure с](https://portal.azure.com) помощью **учетных данных Microsoft 365 администратора**.

1. Перейдите на **страницу Регистрация приложений** и выберите регистрацию приложения, созданную во время быстрого запуска.
    > [!TIP]
    > **Отображаемое имя** приложения соответствует имени надстройки, указанному при создании проекта с помощью генератора Yeoman.

1. В **разделе "Управление**" выберите **разрешения API**.

1. В **строке User.Read** таблицы разрешений нажмите кнопку с многоточием, а  затем в появившемся меню выберите "Отозвать согласие администратора".

    :::image type="content" source="../images/app-registration-revoke-admin-consent.png" alt-text="Снимок экрана: кнопка &quot;Отозвать согласие администратора&quot; на странице разрешений API.":::

1. Нажмите **кнопку "Да",** "Удалить" в ответ на отображаемый запрос.

1. В **строке User.Read** таблицы разрешений нажмите кнопку с многоточием и выберите  "Удалить разрешение" в появившемся меню.

    :::image type="content" source="../images/app-registration-remove-permission.png" alt-text="Снимок экрана: кнопка &quot;Удалить разрешение&quot; на странице разрешений API.":::

1. Нажмите **кнопку "Да",** "Удалить" в ответ на отображаемый запрос.

1. Нажмите кнопку **Добавить разрешение**.

1. На открываемой панели выберите **Microsoft Graph** а затем выберите **делегированные разрешения**.

1. На панели **разрешений API запросов** :

    а. В **разделе "Файлы**" выберите **Files.Read.All**.

    б. Нажмите **кнопку "Добавить разрешения** " в нижней части панели, чтобы сохранить эти изменения разрешений.

1. Нажмите **кнопку "Предоставить согласие администратора для [имя клиента** ]".

1. Нажмите **кнопку "** Да" в ответ на отображаемый запрос.

### <a name="update-code-in-the-add-in-project"></a>Обновление кода в проекте надстройки

Чтобы надстройка считыла содержимое учетной записи пользователя, выполнившего вход, OneDrive для бизнеса:

- Обновите код, который ссылается на URL-Graph Microsoft, параметры и требуемую область доступа.

- Обновите код, определяющий пользовательский интерфейс области задач, чтобы он точно описывал новые функциональные возможности.

- Обновите код, который анализирует ответ от Microsoft Graph и записывает его в документ или сообщение.

Эти обновления описаны в следующих шагах.

### <a name="changes-required-for-any-type-of-add-in"></a>Изменения, необходимые для надстройки любого типа

Выполните следующие действия для надстройки, чтобы изменить URL-Graph Microsoft Graph, параметры и область доступа, а также обновить пользовательский интерфейс области задач. Эти действия одинаковы независимо от того, Office приложения, целевые объекты надстройки.

1. В **./. ENV-файл** :

    а. Заменить `GRAPH_URL_SEGMENT=/me` на `GRAPH_URL_SEGMENT=/me/drive/root/children`

    б. Заменить `QUERY_PARAM_SEGMENT=` на `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    c. Заменить `SCOPE=User.Read` на `SCOPE=Files.Read.All`

1. В **./manifest.xml** найдите `<Scope>User.Read</Scope>` строку в конце файла и замените ее строкой `<Scope>Files.Read.All</Scope>`.

1. В **файле ./src/helpers/fallbackauthdialog.js** ( **или в файле ./src/helpers/fallbackauthdialog.ts** для проекта TypeScript) `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All`найдите строку и замените ее строкой, `requestObj` которая определена следующим образом:

    ```javascript
    var requestObj = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

    ```typescript
    var requestObj: Object = {
      scopes: [`https://graph.microsoft.com/Files.Read.All`]
    };
    ```

1. В **./src/taskpane/taskpane.html**`<section class="ms-firstrun-instructionstep__header">` найдите элемент и обновите текст в этом элементе, чтобы описать новые функции надстройки.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

1. В **./src/taskpane/taskpane.html** найдите `Get My User Profile Information` оба вхождения строки и замените ее `Read my OneDrive for Business`на .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">Click the <b>Read my OneDrive for Business</b>
            button.</span>
        <div class="clearfix"></div>
    </li>
    ```

    ```html
    <p align="center">
        <button id="getGraphDataButton" class="popupButton ms-Button ms-Button--primary"><span
                class="ms-Button-label">Read my OneDrive for Business</span></button>
    </p>
    ```

1. В **./src/taskpane/taskpane.html** найдите строку `Your user profile information will be displayed in the document.` и замените ее `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`на .

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

1. Обновите код, который анализирует ответ от Microsoft Graph и записывает его в документ или сообщение, следуя указаниям в разделе, соответствующем типу надстройки:

    - [Изменения, необходимые для надстройки Office (JavaScript)](#changes-required-for-an-office-add-in-javascript)
    - [Изменения, необходимые для надстройки Office (TypeScript)](#changes-required-for-an-office-add-in-typescript)

### <a name="changes-required-for-an-office-add-in-javascript"></a>Изменения, необходимые для надстройки Office (JavaScript)

Если созданная Office использует JavaScript, внесите следующие изменения в **файл ./src/helpers/documentHelper.js**.

1. Найдите функцию `filterUserProfileInfo` и замените ее следующей функцией.

    ```javascript
    function filterOneDriveInfo(result) {
      let itemNames = [];
      let oneDriveItems = result['value'];
      for (let item of oneDriveItems) {
        itemNames.push(item['name']);
      }
      return itemNames;
    }
    ```

1. Найдите и `filterUserProfileInfo` замените его .`filterOneDriveInfo` Для замены должно быть четыре экземпляра.

1. Сохраните изменения.

После внесения этих изменений перейдите к разделу "Попробовать" [](#try-it-out) этой статьи, чтобы опробовать обновленную надстройку.

### <a name="changes-required-for-an-office-add-in-typescript"></a>Изменения, необходимые для надстройки Office (TypeScript)

Если созданная Office использует TypeScript, откройте **файл ./src/taskpane/taskpane.ts**.

1. Найдите `writeDataToOfficeDocument` функцию и замените ее приведенным ниже кодом в зависимости Office, где будет размещена надстройка (Excel, Outlook, Word или PowerPoint).

#### <a name="excel-code"></a>Excel кода

```typescript
  export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[][];

    // Get just the filenames from results
    data = result["value"].map((item) => {
      return [item.name];
    });

    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

#### <a name="outlook-code"></a>Outlook кода

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });

  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "</br>";
  }
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}
```

#### <a name="word-code"></a>Код Word

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function (context) {
    // Get just the filenames from results.
    const data: string[] = result["value"].map((item) => {
      return item.name;
    });

    const documentBody: Word.Body = context.document.body;
    for (let i = 0; i < data.length; i++) {
      if (data[i] !== null) {
        documentBody.insertParagraph(data[i], "End");
      }
    }
    return context.sync();
  });
}
```

#### <a name="powerpoint-code"></a>PowerPoint кода

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  // Get just the filenames from results.
  const data: string[] = result["value"].map((item) => {
    return item.name;
  });
  let userInfo: string = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(userInfo, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

## <a name="try-it-out"></a>Проверка

Если надстройка является Excel, Word или PowerPoint надстройка, выполните действия, описанные в следующем разделе, чтобы опробовать ее. Если надстройка является Outlook надстройка, выполните действия, описанные в [Outlook разделе.](#outlook)

### <a name="excel-word-and-powerpoint"></a>Excel, Word и PowerPoint

Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.

1. В корневой папке проекта выполните следующую команду, чтобы выполнить сборку проекта, запустить локальный веб-сервер и загрузить неопубликованную надстройку в ранее выбранном Office клиентском приложении.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. В клиентском приложении Office, которое открывается при выполнении предыдущей команды (например, Excel, Word или PowerPoint), убедитесь, что вы вошли с пользователем, который является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которая использовалась для подключения к Azure при настройке единого входа.[](sso-quickstart.md#configure-sso)  для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа. 

3. В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки. На рисунке ниже показана эта кнопка в Excel.

    ![Снимок экрана: выделенная кнопка надстройки на Excel ленте.](../images/excel-quickstart-addin-3b.png)

4. В нижней части области задач нажмите кнопку "Чтение **OneDrive для бизнеса",** чтобы инициировать процесс единого входа.

5. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана диалогового окна, запрашивающего разрешение, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

6. Надстройка считывает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает в документ имена 10 основных файлов и папок. На следующем рисунке показан пример имен файлов и папок, записанных на Excel листа.

    ![Снимок экрана: OneDrive для бизнеса сведения на Excel листе.](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Выполните следующие действия, чтобы испытать надстройку Outlook.

1. В корневой папке проекта выполните следующую команду, чтобы создать проект, запустить локальный веб-сервер и загрузить неопубликованную надстройку. 

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    ```command&nbsp;line
    npm start
    ```

2. Убедитесь, что вы вошли в Outlook с пользователем, который является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которая использовалась для подключения к Azure при настройке единого входа для приложения.[](sso-quickstart.md#configure-sso) Благодаря этому будут созданы соответствующие условия для успешного единого входа.

3. В Outlook создайте новое сообщение.

4. В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: выделенная кнопка ленты надстройки в окне создания сообщения Outlook.](../images/outlook-sso-ribbon-button.png)

5. В нижней части области задач нажмите кнопку "Чтение **OneDrive для бизнеса",** чтобы инициировать процесс единого входа.

6. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана: диалоговое окно, запрашивающее разрешения, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

7. Надстройка считывает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает имена 10 основных файлов и папок в текст сообщения электронной почты.

    ![Снимок экрана: OneDrive для бизнеса сведения в Outlook создания сообщения.](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно настроите функциональность надстройки с поддержкой единого входа, созданной с помощью генератора Yeoman в кратком руководстве по [единому входу](sso-quickstart.md). Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>См. также

- [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md)
- [Краткое руководство по единому входу (SSO)](sso-quickstart.md)
- [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md)
- [Устранение ошибок единого входа](../develop/troubleshoot-sso-in-office-add-ins.md)
