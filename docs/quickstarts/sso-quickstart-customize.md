---
title: Настройка надстройки Node.js с поддержкой единого входа
description: Узнайте о настройке надстройки с поддержкой SSO, созданной с помощью генератора Yeoman.
ms.date: 08/04/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 8b66da764902fba1a0296c349f898a94ef9f9b33
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154554"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Настройка надстройки Node.js с поддержкой единого входа

> [!IMPORTANT]
> В этой статье построят надстройку с поддержкой SSO, созданную путем быстрого запуска единого входа [(SSO).](sso-quickstart.md) Перед чтением этой статьи выполните быстрое начало.

Быстрое начало [SSO](sso-quickstart.md) создает надстройки с поддержкой SSO, которая получает сведения о профиле пользователя и записывает их в документ или сообщение. В этой статье вы пройдите процесс обновления надстройки, созданной с помощью генератора Yeoman в быстром запуске SSO, чтобы добавить новые функциональные возможности, которые требуют различных разрешений.

## <a name="prerequisites"></a>Предварительные требования

- Надстройка Office, которую вы создали, следуя инструкциям в [быстром запуске SSO.](sso-quickstart.md)

- По крайней мере несколько файлов и папок, OneDrive для бизнеса в Microsoft 365 подписке.

- [Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases)).

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Просмотр содержимого проекта

Начнем с краткого обзора проекта надстройки, созданного ранее с помощью генератора [Yeoman.](sso-quickstart.md)

> [!NOTE]
> В местах, где в  этой статье ссылаются файлы сценариев с.jsрасширения файлов, предположим расширение **файла .ts,** если проект был создан с помощью TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Добавление новых функций

Надстройка, созданная с помощью быстрого запуска SSO, использует microsoft Graph для получения сведений о профиле пользователя и записывает эти сведения в документ или сообщение. Давайте изменим функции надстройки таким образом, чтобы она получила имена 10 лучших файлов и папок из записи пользователя с OneDrive для бизнеса и записывает эти сведения в документ или сообщение. Включение этой новой функции требует обновления разрешений приложений в Azure и обновления кода в проекте надстройки.

### <a name="update-app-permissions-in-azure"></a>Обновление разрешений приложений в Azure

Прежде чем надстройка сможет успешно прочитать содержимое OneDrive для бизнеса пользователя, сведения о регистрации приложений в Azure должны обновляться с соответствующими разрешениями. Выполните следующие действия, чтобы предоставить приложению разрешение **Files.Read.All** и отопросить разрешение **User.Read,** которое больше не требуется.

1. Перейдите на [портал Azure и](https://ms.portal.azure.com/#home) войдите с помощью учетных данных Microsoft 365 **администратора.**

2. Перейдите на **страницу регистрации приложений.**
    > [!TIP]
    > Это можно сделать, выбрав  плитку регистрации приложений на домашней странице Azure или используя поле поиска на домашней странице, чтобы найти и выбрать регистрации **приложений.**

3. На странице **Регистрация приложений** выберите приложение, созданное во время быстрого запуска.
    > [!TIP]
    > Имя **отображения приложения** будет соответствовать имени надстройки, указанному при создания проекта с генератором Yeoman.

4. На странице обзор приложения выберите **разрешения API** в заголовке **Управление** в левой части страницы.

5. В **строке User.Read** таблицы разрешений выберите ellipsis, а затем выберите согласие администратора отопросить из меню, которое отображается. 

6. Выберите **кнопку Да, удалите** кнопку в ответ на отображаемую подсказку.

7. В **строке User.Read** таблицы разрешений выберите ellipsis  и выберите Удаление разрешения из меню, которое отображается.

8. Выберите **кнопку Да, удалите** кнопку в ответ на отображаемую подсказку.

9. Нажмите кнопку **Добавить разрешение**.

10. На открываемой панели выберите **Microsoft Graph** а затем выберите **делегированную разрешения.**

11. На панели **разрешений API запроса:**

    а. В **статье Файлы** выберите **Files.Read.All**.

    б. Выберите **кнопку Добавить разрешения** в нижней части панели, чтобы сохранить эти изменения разрешений.

12. Выберите согласие **администратора гранта для кнопки [имя клиента].**

13. Выберите **кнопку Да** в ответ на отображаемую подсказку.

### <a name="update-code-in-the-add-in-project"></a>Обновление кода в проекте надстройки

Чтобы надстройка считывала содержимое пользовательского OneDrive для бизнеса, необходимо:

- Обнови код, который ссылается на URL Graph Microsoft, параметры и область требуемого доступа.

- Обнови код, определяя пользовательский интерфейс области задач, чтобы он точно описывал новые функции.

- Обновите код, который размазирует ответ от Microsoft Graph и нанося его в документ или сообщение.

Ниже описаны эти обновления.

### <a name="changes-required-for-any-type-of-add-in"></a>Изменения, необходимые для любого типа надстройки

Выполните следующие действия для надстройки, чтобы изменить URL Graph Microsoft, параметры и область доступа, а также обновить пользовательский интерфейс области задач. Эти действия одинаковы, независимо от того, Office приложения ваших целей надстройки.

1. В **./. Файл ENV:**

    а. Замените `GRAPH_URL_SEGMENT=/me` следующим образом: `GRAPH_URL_SEGMENT=/me/drive/root/children`

    б. Замените `QUERY_PARAM_SEGMENT=` следующим образом: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    в. Замените `SCOPE=User.Read` следующим образом: `SCOPE=Files.Read.All`

2. В **./manifest.xml** найти строку в конце файла и заменить `<Scope>User.Read</Scope>` ее строкой `<Scope>Files.Read.All</Scope>` .

3. В **./src/helpers/fallbackauthdialog.js** (или **в ./src/helpers/fallbackauthdialog.ts** для проекта TypeScript) найдите строку и замените ее строкой, которая определяется следующим `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` образом:

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

4. В **./src/taskpane/taskpane.html** найти элемент и обновить текст в этом элементе, чтобы описать новые функциональные `<section class="ms-firstrun-instructionstep__header">` возможности надстройки.

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. В **./src/taskpane/taskpane.html** найти и заменить оба появления `Get My User Profile Information` строки `Read my OneDrive for Business` строкой.

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

6. В **./src/taskpane/taskpane.html** найти и заменить `Your user profile information will be displayed in the document.` строку `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` строкой.

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Обновите код, который размывает ответ от Microsoft Graph и записывает его в документ или сообщение, следуя указаниям в разделе, соответствующем типу надстройки:

    - [Изменения, необходимые для Excel надстройки (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Изменения, необходимые для Excel надстройки (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Изменения, необходимые для Outlook надстройки (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Изменения, необходимые для Outlook надстройки (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [Изменения, необходимые для PowerPoint надстройки (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [Изменения, необходимые для PowerPoint надстройки (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Изменения, необходимые для надстройки Word (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Изменения, необходимые для надстройки Word (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Изменения, необходимые для Excel надстройки (JavaScript)

Если надстройка является Excel, созданной с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js.**

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией.

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

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией.

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

3. Найдите `writeDataToExcel` функцию и замените ее следующей функцией.

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

4. Удаление `writeDataToOutlook` функции.

5. Удаление `writeDataToPowerPoint` функции.

6. Удаление `writeDataToWord` функции.

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Изменения, необходимые для Excel надстройки (TypeScript)

Если надстройка является Excel, созданной с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией.

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

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Изменения, необходимые для Outlook надстройки (JavaScript)

Если надстройка Outlook, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js.**

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией.

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

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией.

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

3. Найдите `writeDataToOutlook` функцию и замените ее следующей функцией.

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

4. Удаление `writeDataToExcel` функции.

5. Удаление `writeDataToPowerPoint` функции.

6. Удаление `writeDataToWord` функции.

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Изменения, необходимые для Outlook надстройки (TypeScript)

Если надстройка Outlook, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией.

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

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>Изменения, необходимые для PowerPoint надстройки (JavaScript)

Если надстройка — это надстройка PowerPoint, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js.**

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией.

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

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией.

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

3. Найдите `writeDataToPowerPoint` функцию и замените ее следующей функцией.

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

4. Удаление `writeDataToExcel` функции.

5. Удаление `writeDataToOutlook` функции.

6. Удаление `writeDataToWord` функции.

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>Изменения, необходимые для PowerPoint надстройки (TypeScript)

Если надстройка — это надстройка PowerPoint, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts**, найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией.

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

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-a-word-add-in-javascript"></a>Изменения, необходимые для надстройки Word (JavaScript)

Если надстройка — это надстройка Word, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js.**

1. Найдите `writeDataToOfficeDocument` функцию и замените ее следующей функцией.

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

2. Найдите `filterUserProfileInfo` функцию и замените ее следующей функцией.

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

3. Найдите `writeDataToWord` функцию и замените ее следующей функцией.

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

4. Удаление `writeDataToExcel` функции.

5. Удаление `writeDataToOutlook` функции.

6. Удаление `writeDataToPowerPoint` функции.

После внесения этих изменений перескочить в раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

### <a name="changes-required-for-a-word-add-in-typescript"></a>Изменения, необходимые для надстройки Word (TypeScript)

Если ваша надстройка — это надстройка Word, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts**, найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией.

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

После внесения этих изменений перенастройте раздел [Try it out](#try-it-out) из этой статьи, чтобы опробуете обновленную надстройка.

## <a name="try-it-out"></a>Проверка

Если надстройка является Excel, Word или PowerPoint надстройка, выполните действия в следующем разделе, чтобы попробовать его. Если надстройка является Outlook надстройка, выполните шаги в [Outlook](#outlook) разделе.

### <a name="excel-word-and-powerpoint"></a>Excel, Word и PowerPoint

Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.

1. В корневой папке проекта запустите следующую команду для создания проекта, запустите локальный веб-сервер и разгрузите надстройку в выбранном ранее клиентом приложении Office.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. В клиентской Office, открываемой при запуске предыдущей команды (например, Excel, Word или PowerPoint), убедитесь, что вы подписаны с пользователем, который является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которую вы использовали для подключения к Azure при настройке [SSO](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа. 

3. В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки. На рисунке ниже показана эта кнопка в Excel. 

    ![Снимок экрана, показывающий выделенную кнопку надстройки в Excel ленте.](../images/excel-quickstart-addin-3b.png)

4. В нижней части области задач выберите кнопку Read **my OneDrive для бизнеса,** чтобы инициировать процесс SSO.

5. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана диалогового окна, запрашивающего разрешение, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

6. Надстройка считывает данные из пользовательского OneDrive для бизнеса и записывает в документ имена 10 лучших файлов и папок. На следующем изображении показан пример имен файлов и папок, написанных на Excel таблицу.

    ![Снимок экрана, OneDrive для бизнеса сведения в Excel таблицы.](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Выполните следующие действия, чтобы испытать надстройку Outlook.

1. В корневой папке проекта запустите следующую команду для создания проекта, запустите локальный веб-сервер и разгрузите надстройку. 

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman. Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.

    ```command&nbsp;line
    npm start
    ```

2. Убедитесь, что вы подписались Outlook с пользователем, который является членом той же Microsoft 365 организации, что и учетная запись Microsoft 365 администратора, которую вы использовали для подключения к Azure при настройке [SSO](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа.

3. В Outlook создайте новое сообщение.

4. В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: выделенная кнопка ленты надстройки в окне создания сообщения Outlook.](../images/outlook-sso-ribbon-button.png)

5. В нижней части области задач выберите кнопку Read **my OneDrive для бизнеса,** чтобы инициировать процесс SSO.

6. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана: диалоговое окно, запрашивающее разрешения, с выделенной кнопкой "Принять".](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

7. Надстройка считывает данные из пользовательского OneDrive для бизнеса и записывает имена 10 лучших файлов и папок в текст сообщения электронной почты.

    ![Снимок экрана, OneDrive для бизнеса сведения в Outlook окне составить сообщение.](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Следующие шаги

Поздравляем, вы успешно настраивали функции надстройки с поддержкой SSO, созданной с генератором Yeoman в быстром запуске [SSO.](sso-quickstart.md) Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>См. также

- [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md)
- [Краткое руководство по единому входу (SSO)](sso-quickstart.md)
- [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md)
- [Устранение ошибок единого входа](../develop/troubleshoot-sso-in-office-add-ins.md)
