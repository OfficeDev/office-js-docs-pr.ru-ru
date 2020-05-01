---
title: Настройка надстройки Node.js с поддержкой единого входа
description: Сведения о настройке надстройки с поддержкой единого входа, созданной с помощью генератора Yeoman.
ms.date: 02/20/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c02e0f74a8ea3f3f8f831b65aa403ce49655953b
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284169"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Настройка надстройки Node.js с поддержкой единого входа

> [!IMPORTANT]
> Эта статья основана на надстройке с поддержкой единого входа, которая создается с помощью краткого руководства по выполнению [единого входа (SSO)](sso-quickstart.md). Прежде чем приступить к чтению этой статьи, заполните краткое руководство.

[Быстрый запуск единого входа](sso-quickstart.md) создает надстройку с включенной поддержкой единого входа, которая получает данные профиля пользователя, выполнившего вход, и записывает их в документ или сообщение. В этой статье описывается процесс обновления надстройки, созданной с помощью генератора Yeoman в быстром запуске единого входа, для добавления новых функциональных возможностей, требующих других разрешений.

## <a name="prerequisites"></a>Необходимые компоненты

* Надстройка Office, созданная в соответствии с инструкциями, приведенными в [кратком](sso-quickstart.md)руководстве по SSO.

* Несколько файлов и папок, сохраненных в OneDrive для бизнеса в составе подписки на Office 365.

* [Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases)).

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Просмотр содержимого проекта

Начнем с краткого обзора проекта надстройки, [созданного ранее с помощью генератора Yeoman](sso-quickstart.md).

> [!NOTE]
> В местах, где эта статья ссылается на файлы сценариев с использованием расширения **JS** , вместо этого следует использовать расширение **TS** , если проект был создан с помощью TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Добавление новых функциональных возможностей 

Надстройка, созданная с помощью быстрого запуска единого входа, использует Microsoft Graph для получения сведений о профиле пользователя, выполнившего вход, и записывает эти сведения в документ или сообщение. Теперь изменим функциональные возможности надстройки, чтобы она выводила имена 10 файлов и папок из OneDrive для бизнеса пользователя, выполнившего вход, и записывает эти сведения в документ или сообщение. Для этого требуется обновление разрешений приложений в Azure и обновление кода в проекте надстройки.

### <a name="update-app-permissions-in-azure"></a>Обновление разрешений приложения в Azure

Прежде чем надстройка сможет успешно прочитать содержимое OneDrive для бизнеса пользователя, ее регистрационная информация в Azure должна быть обновлена с соответствующими разрешениями. Выполните следующие действия, чтобы предоставить приложению разрешение **Files. Read. ALL** и отозвать разрешение **User.** Read. ALL, что больше не требуется.

1. Перейдите на [портал Azure](https://ms.portal.azure.com/#home) и **Войдите в систему, используя учетные данные администратора Office 365**. 

2. Перейдите на страницу **регистрации приложений** . 
    > [!TIP]
    > Это можно сделать, выбрав плитку **регистрации приложений** на домашней странице Azure или воспользовавшись полем поиска на домашней странице, чтобы найти и выбрать **регистрации приложений**.

3. На странице **регистрации приложений** выберите приложение, созданное на этапе быстрого запуска. 
    > [!TIP]
    > **Отображаемое имя** приложения будет соответствующим имени надстройки, которое вы указали при создании проекта с помощью генератора Yeoman.

4. На странице "Обзор приложения" выберите **разрешения API** в разделе **Управление** заголовком в левой части страницы.

5. В строке **User. Read** таблицы Permissions нажмите кнопку с многоточием, а затем выберите **отозвать согласие администратора** из появившегося меню.

6. Нажмите кнопку **Да, удалить** в ответ на отображаемый запрос.

7. В строке **User. Read** таблицы Permissions нажмите кнопку с многоточием, а затем выберите пункт **удалить разрешение** из появившегося меню.

8. Нажмите кнопку **Да, удалить** в ответ на отображаемый запрос.

9. Нажмите кнопку **Добавить разрешение** .

10. В открывшейся панели выберите **Microsoft Graph** , а затем — **делегированные разрешения**.

11. На панели **разрешений API запроса** выполните следующие действия:

    а. В разделе **файлы**выберите **файлы. Read. ALL**.

    б) Нажмите кнопку **Добавить разрешения** в нижней части панели, чтобы сохранить изменения этих разрешений.

12. Нажмите кнопку **предоставить согласие администратора для пользователя [имя клиента]** .

13. Нажмите кнопку **Да** в ответ на отображаемый запрос.

### <a name="update-code-in-the-add-in-project"></a>Обновление кода в проекте надстройки

Чтобы надстройка прочитала содержимое OneDrive для бизнеса пользователя, выполнившего вход, необходимо выполнить следующие действия:

- Обновите код, ссылающийся на URL-адрес, параметры и требуемую область доступа Microsoft Graph.

- Обновите код, определяющий пользовательский интерфейс области задач, чтобы он точно описывает новые функциональные возможности. 

- Обновление кода, который анализирует отклик от Microsoft Graph и записывает его в документ или сообщение.

Эти обновления описываются в следующих шагах.

### <a name="changes-required-for-any-type-of-add-in"></a>Изменения, необходимые для любого типа надстройки

Выполните указанные ниже действия для надстройки, чтобы изменить URL-адрес, параметры и область доступа Microsoft Graph, а также обновить пользовательский интерфейс области задач. Эти действия одинаковы, независимо от того, в каком приложении Office размещены целевые объекты надстройки.

1. В файле **./. ENV** :

    а. Замените `GRAPH_URL_SEGMENT=/me` на следующий:`GRAPH_URL_SEGMENT=/me/drive/root/children`

    б) Замените `QUERY_PARAM_SEGMENT=` на следующий:`QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    в. Замените `SCOPE=User.Read` на следующий:`SCOPE=Files.Read.All`

2. В **/манифест.ксмл**найдите строку `<Scope>User.Read</Scope>` около конца файла и замените ее на строку `<Scope>Files.Read.All</Scope>`.

3. В файле **./СРК/Хелперс/фаллбаккаусдиалог.ЖС** (или **в/СРК/Хелперс/фаллбаккаусдиалог.ТС** для проекта TypeScript) найдите `https://graph.microsoft.com/User.Read` строку и замените ее строкой `https://graph.microsoft.com/Files.Read.All`, которая `requestObj` определяется следующим образом:

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

4. В **src/TaskPane/TaskPane.HTML**найдите элемент `<section class="ms-firstrun-instructionstep__header">` и обновите текст в этом элементе, чтобы описать новые функции надстройки.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. В файле **./src/TaskPane/TaskPane.HTML**найдите и замените все вхождения строки `Get My User Profile Information` строкой. `Read my OneDrive for Business`

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

6. В файле **./src/TaskPane/TaskPane.HTML**найдите и замените строку `Your user profile information will be displayed in the document.` строкой. `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.`

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Обновите код, который анализирует ответ от Microsoft Graph, и записывает его в документ или сообщение, следуя указаниям в разделе, соответствующем типу надстройки:

    - [Изменения, необходимые для надстройки Excel (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Изменения, необходимые для надстройки Excel (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Изменения, необходимые для надстройки Outlook (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Изменения, необходимые для надстройки Outlook (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [Изменения, необходимые для надстройки PowerPoint (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [Изменения, необходимые для надстройки PowerPoint (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Изменения, необходимые для надстройки Word (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Изменения, необходимые для надстройки Word (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Изменения, необходимые для надстройки Excel (JavaScript)

Если надстройка представляет собой надстройку Excel, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToExcel(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:

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

3. Найдите `writeDataToExcel` функцию и замените ее следующей функцией:

    ```javascript
    function writeDataToExcel(result) {
      return Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getActiveWorksheet();
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            let innerArray = [];
            innerArray.push(oneDriveInfo[i]);
            data.push(innerArray);
          }
        }

        const rangeAddress = `B5:B${5 + (data.length - 1)}`;
        const range = sheet.getRange(rangeAddress);
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
      });
    }
    ```

4. Удалите `writeDataToOutlook` функцию.

5. Удалите `writeDataToPowerPoint` функцию.

6. Удалите `writeDataToWord` функцию.

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Изменения, необходимые для надстройки Excel (TypeScript)

Если надстройка представляет собой надстройку Excel, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Excel.run(function(context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    }

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        let innerArray = [];
        innerArray.push(itemNames[i]);
        data.push(innerArray);
      }
    }
    
    const rangeAddress = `B5:B${5 + (data.length - 1)}`;
    const range = sheet.getRange(rangeAddress);
    range.values = data;
    range.format.autofitColumns();

    return context.sync();
  });
}
```

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Изменения, необходимые для надстройки Outlook (JavaScript)

Если надстройка представляет собой надстройку Outlook, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToOutlook(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to message. " + error.toString()));
        }
      });
    }
    ```

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:

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

3. Найдите `writeDataToOutlook` функцию и замените ее следующей функцией:

    ```javascript
    function writeDataToOutlook(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
      }

      Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
    }
    ```

4. Удалите `writeDataToExcel` функцию.

5. Удалите `writeDataToPowerPoint` функцию.

6. Удалите `writeDataToWord` функцию.

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Изменения, необходимые для надстройки Outlook (TypeScript)

Если надстройка представляет собой надстройку Outlook, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

```typescript
export function writeDataToOfficeDocument(result: Object): void {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
        itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
        if (itemNames[i] !== null) {
        data.push(itemNames[i]);
        }
    }

    let objectNames: string = "";
    for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "<br/>";
    }
    
    Office.context.mailbox.item.body.setSelectedDataAsync(objectNames, { coercionType: Office.CoercionType.Html });
}
```

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>Изменения, необходимые для надстройки PowerPoint (JavaScript)

Если надстройка представляет собой надстройку PowerPoint, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToPowerPoint(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:

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

3. Найдите `writeDataToPowerPoint` функцию и замените ее следующей функцией:

    ```javascript
    function writeDataToPowerPoint(result) {
      let data = [];
      let oneDriveInfo = filterOneDriveInfo(result);

      for (let i = 0; i < oneDriveInfo.length; i++) {
        if (oneDriveInfo[i] !== null) {
          data.push(oneDriveInfo[i]);
        }
      }

      let objectNames = "";
      for (let i = 0; i < data.length; i++) {
        objectNames += data[i] + "\n";
      }

      Office.context.document.setSelectedDataAsync(
        objectNames, 
        function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            throw asyncResult.error.message;
          }
      });
    }
    ```

4. Удалите `writeDataToExcel` функцию.

5. Удалите `writeDataToOutlook` функцию.

6. Удалите `writeDataToWord` функцию.

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>Изменения, необходимые для надстройки PowerPoint (TypeScript)

Если надстройка представляет собой надстройку PowerPoint, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

```typescript
export function writeDataToOfficeDocument(result: Object): void {
  let data: string[] = [];

  let itemNames: string[] = [];
  let oneDriveItems = result["value"];
  for (let item of oneDriveItems) {
    itemNames.push(item["name"]);
  };

  for (let i = 0; i < itemNames.length; i++) {
    if (itemNames[i] !== null) {
      data.push(itemNames[i]);
    }
  }

  let objectNames: string = "";
  for (let i = 0; i < data.length; i++) {
    objectNames += data[i] + "\n";
  }

  Office.context.document.setSelectedDataAsync(objectNames, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      throw asyncResult.error.message;
    }
  });
}
```

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-a-word-add-in-javascript"></a>Изменения, необходимые для надстройки Word (JavaScript)

Если надстройка представляет собой надстройку Word, созданную с помощью JavaScript, внесите следующие изменения в файле **./СРК/Хелперс/докуменселпер.ЖС**:

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

    ```javascript
    export function writeDataToOfficeDocument(result) {
      return new OfficeExtension.Promise(function(resolve, reject) {
        try {
          writeDataToWord(result);
          resolve();
        } catch (error) {
          reject(Error("Unable to write data to document. " + error.toString()));
        }
      });
    }
    ```

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией:

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

3. Найдите `writeDataToWord` функцию и замените ее следующей функцией:

    ```javascript
    function writeDataToWord(result) {
      return Word.run(function (context) {
        let data = [];
        let oneDriveInfo = filterOneDriveInfo(result);

        for (let i = 0; i < oneDriveInfo.length; i++) {
          if (oneDriveInfo[i] !== null) {
            data.push(oneDriveInfo[i]);
          }
        }

        const documentBody = context.document.body;
        for (let i = 0; i < data.length; i++) {
          if (data[i] !== null) {
            documentBody.insertParagraph(data[i], "End");
          }
        }

        return context.sync();
      });
    }
    ```

4. Удалите `writeDataToExcel` функцию.

5. Удалите `writeDataToOutlook` функцию.

6. Удалите `writeDataToPowerPoint` функцию.

После внесения этих изменений перейдите к разделу " [попробовать](#try-it-out) " в этой статье, чтобы испытать обновленную надстройку.

### <a name="changes-required-for-a-word-add-in-typescript"></a>Изменения, необходимые для надстройки Word (TypeScript)

Если надстройка представляет собой надстройку Word, созданную с помощью TypeScript, откройте **./СРК/таскпане/таскпане.ТС**, найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией:

```typescript
export function writeDataToOfficeDocument(result: Object): Promise<any> {
  return Word.run(function(context) {
    let data: string[] = [];

    let itemNames: string[] = [];
    let oneDriveItems = result["value"];
    for (let item of oneDriveItems) {
      itemNames.push(item["name"]);
    };

    for (let i = 0; i < itemNames.length; i++) {
      if (itemNames[i] !== null) {
        data.push(itemNames[i]);
      }
    }

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

После внесения этих изменений перейдите [к разделу](#try-it-out) "ознакомьтесь с этой статьей", чтобы испытать обновленную надстройку.

## <a name="try-it-out"></a>Проверка

Если надстройка представляет собой надстройку Excel, Word или PowerPoint, выполните действия, описанные в следующем разделе, чтобы попробовать. Если надстройка является надстройкой Outlook, выполните действия, описанные в разделе [Outlook](#outlook) .

### <a name="excel-word-and-powerpoint"></a>Excel, Word и PowerPoint

Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.

1. В корневой папке проекта выполните следующую команду, чтобы выполнить сборку проекта, запустите локальный веб-сервер и Загрузка неопубликованных вашу надстройку в выбранном ранее клиентском приложении Office.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. В клиентском приложении Office, которое открывается при выполнении предыдущей команды (например, Excel, Word или PowerPoint), убедитесь, что вы вошли в систему с учетной записью пользователя, который является участником той же организации Office 365, что и учетная запись администратора Office 365, которую вы использовали для подключения к Azure при [настройке единого входа](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа. 

3. В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки. На рисунке ниже показана эта кнопка в Excel. 

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. В нижней части области задач нажмите кнопку **прочитать мою службу OneDrive для бизнеса** , чтобы начать процесс единого входа. 

5. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи). Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

6. Надстройка читает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает в документ имена из 10 самых популярных файлов и папок. На следующем рисунке показан пример имен файлов и папок, записанных на лист Excel.

    ![Сведения о OneDrive для бизнеса в таблице Excel](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Выполните следующие действия, чтобы испытать надстройку Outlook.

1. В корневой папке проекта выполните следующую команду, чтобы построить проект и запустить локальный веб-сервер.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. Чтобы загрузить неопубликованную надстройку в Outlook, следуйте инструкциями из статьи [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing). Убедитесь, что вы выполнили вход в Outlook с пользователем, который является участником той же организации Office 365, что и учетная запись администратора Office 365, которую вы использовали для подключения к Azure при [настройке единого входа](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа. 

3. В Outlook создайте новое сообщение.

4. В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Outlook](../images/outlook-sso-ribbon-button.png)

5. В нижней части области задач нажмите кнопку **прочитать мою службу OneDrive для бизнеса** , чтобы начать процесс единого входа. 

6. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или Office 365 (рабочей или учебной учетной записи). Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Диалоговое окно запроса разрешений](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

7. Надстройка читает данные из OneDrive для бизнеса пользователя, выполнившего вход, и записывает имена 10 файлов и папок в текст сообщения электронной почты.

    ![Сведения о OneDrive для бизнеса в сообщении Outlook](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно настроили функции надстройки с поддержкой единого входа, созданной с помощью генератора Yeoman в [быстром запуске единого входа](sso-quickstart.md). Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>См. также

- [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md)
- [Краткое руководство по единому входу (SSO)](sso-quickstart.md)
- [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md)
- [Устранение ошибок единого входа](../develop/troubleshoot-sso-in-office-add-ins.md)
