---
title: Настройка надстройки Node.js с поддержкой единого входа
description: Узнайте о настройке надстройки с поддержкой SSO, созданной с помощью генератора Yeoman.
ms.date: 02/01/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 174df5e58e794b94b02025bd90a65f5ae8e26d44
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234172"
---
# <a name="customize-your-nodejs-sso-enabled-add-in"></a>Настройка надстройки Node.js с поддержкой единого входа

> [!IMPORTANT]
> Эта статья создана на основе надстройки с поддержкой единого входа, созданной путем быстрого запуска [единого входа.](sso-quickstart.md) Прежде чем прочитать эту статью, выполните краткое начало работы.

В кратком начале [SSO](sso-quickstart.md) создается надстройка с поддержкой SSO, которая получает данные профиля во включенного пользователя и записывает их в документ или сообщение. В этой статье вы разберемся в процессе обновления надстройки, созданной с помощью генератора Yeoman в кратком запуске SSO, чтобы добавить новые функции, требуя различных разрешений.

## <a name="prerequisites"></a>Предварительные условия

- Надстройка Office, созданная с помощью инструкций в кратком начале службы [SSO.](sso-quickstart.md)

- По крайней мере несколько файлов и папок, хранимые в OneDrive для бизнеса в вашей подписке на Microsoft 365.

- [Node.js](https://nodejs.org) (последняя версия [LTS](https://nodejs.org/about/releases)).

[!include[additional prerequisites](../includes/sso-tutorial-prereqs.md)]

## <a name="review-contents-of-the-project"></a>Просмотр содержимого проекта

Начнем с краткого просмотра проекта надстройки, который вы ранее создали [с помощью генератора Yeoman.](sso-quickstart.md)

> [!NOTE]
> В тех местах, где эта статья ссылается на файлы скриптов с расширением **JS,** предположим, что расширение **TS** файла, если ваш проект был создан с помощью TypeScript.

[!include[project structure for an SSO-enabled add-in created with the Yeoman generator](../includes/sso-yeoman-project-structure.md)]

## <a name="add-new-functionality"></a>Добавление новых функций

Надстройка, созданная с помощью краткого запуска службы SSO, использует Microsoft Graph для получения сведений о профиле во пользователя и записи этих сведений в документ или сообщение. Давайте изменим функции надстройки так, чтобы она получила имена 10 самых 10 файлов и папок из OneDrive для бизнеса во пользователя, выписав эти сведения в документ или сообщение. Для включения этой новой функции требуется обновление разрешений приложения в Azure и обновление кода в проекте надстройки.

### <a name="update-app-permissions-in-azure"></a>Обновление разрешений приложения в Azure

Прежде чем надстройка сможет успешно прочитать содержимое OneDrive для бизнеса пользователя, сведения о регистрации ее приложений в Azure должны быть обновлены с соответствующими разрешениями. Выполните следующие действия, чтобы предоставить приложению разрешение **Files.Read.All** и отопросить разрешение **User.Read,** которое больше не требуется.

1. Перейдите на [портал Azure и](https://ms.portal.azure.com/#home) войдите, используя учетные данные администратора Microsoft **365.**

2. Перейдите на **страницу регистрации** приложений.
    > [!TIP]
    > Это можно сделать, выбрав  плитку регистрации приложений на домашней странице Azure или используя поле поиска на домашней странице, чтобы найти и выбрать регистрацию **приложений.**

3. На странице **регистрации приложений** выберите приложение, созданное во время краткого запуска.
    > [!TIP]
    > **Отображаемая имя** приложения будет совпадать с именем надстройки, которое вы указали при создания проекта с помощью генератора Yeoman.

4. На странице обзора приложения выберите **разрешения API** под заголовком **"Управление"** в левой части страницы.

5. В **строке User.Read** таблицы разрешений выберите многоточки,  а затем выберите "Отопросить согласие администратора" в отображатом меню.

6. Выберите **кнопку "Да", "Удалить"** в ответ на отображаемую подсказку.

7. В **строке User.Read** таблицы разрешений выберите многоточки,  а затем выберите "Удалить разрешение" в меню.

8. Выберите **кнопку "Да", "Удалить"** в ответ на отображаемую подсказку.

9. Выберите **кнопку "Добавить разрешение".**

10. На открываемой панели выберите **Microsoft Graph,** а затем выберите "Делегирование **разрешений".**

11. На панели **разрешений API запросов:**

    а. Under **Files**, select **Files.Read.All**.

    б. Выберите **кнопку "Добавить разрешения"** в нижней части панели, чтобы сохранить эти изменения разрешений.

12. Выберите **кнопку "Предоставить согласие администратора для [имя клиента]".**

13. Выберите **кнопку "Да"** в ответ на отображаемую подсказку.

### <a name="update-code-in-the-add-in-project"></a>Обновление кода в проекте надстройки

Чтобы надстройка считывала содержимое OneDrive для бизнеса во выгрузаемого пользователя, необходимо:

- Обновим код, который ссылается на URL-адрес Microsoft Graph, параметры и требуемую область доступа.

- Обновите код, который определяет пользовательский интерфейс области задач, чтобы он точно описывал новые функции.

- Обновите код, который проансирует ответ из Microsoft Graph и записывает его в документ или сообщение.

Эти обновления описаны в следующих шагах.

### <a name="changes-required-for-any-type-of-add-in"></a>Изменения, необходимые для любого типа надстройки

Выполните следующие действия для надстройки, чтобы изменить URL-адрес Microsoft Graph, параметры и область доступа, а также обновить пользовательский интерфейс области задач. Эти действия одинаковы независимо от того, какое приложение Office будет целевым для надстройки.

1. В **./. ENV-файл:**

    а. Замените `GRAPH_URL_SEGMENT=/me` следующим образом: `GRAPH_URL_SEGMENT=/me/drive/root/children`

    б. Замените `QUERY_PARAM_SEGMENT=` следующим образом: `QUERY_PARAM_SEGMENT=?$select=name&$top=10`

    в. Замените `SCOPE=User.Read` следующим образом: `SCOPE=Files.Read.All`

2. В **./manifest.xml** найдите строку в конце файла и замените ее `<Scope>User.Read</Scope>` `<Scope>Files.Read.All</Scope>` строкой.

3. В **./src/helpers/fallbackauthdialog.js** (или в **./src/helpers/fallbackauthdialog.ts для** проекта TypeScript) найдите строку и замените ее строкой, которая определяется следующим `https://graph.microsoft.com/User.Read` `https://graph.microsoft.com/Files.Read.All` `requestObj` образом:

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

4. В **./src/taskpane/taskpane.html** найдите элемент и обновите текст в этом элементе, чтобы описать новую функциональность надстройки. `<section class="ms-firstrun-instructionstep__header">`

    ```html
    <section class="ms-firstrun-instructionstep__header">
        <h2 class="ms-font-m">This add-in demonstrates how to use single sign-on by making a call to Microsoft
            Graph to read content from OneDrive for Business.</h2>
        <div class="ms-firstrun-instructionstep__header--image"></div>
    </section>
    ```

5. В **./src/taskpane/taskpane.html** найдите и замените оба вхождения `Get My User Profile Information` строки строкой. `Read my OneDrive for Business`

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

6. В **./src/taskpane/taskpane.html** найдите и замените строку `Your user profile information will be displayed in the document.` `The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.` строкой.

    ```html
    <li class="ms-ListItem">
        <span class="ms-ListItem-primaryText">The names of the top 10 files and folders in your OneDrive for Business will be displayed in the document or message.</span>
        <div class="clearfix"></div>
    </li>
    ```

7. Обновите код, который проансирует ответ из Microsoft Graph и записывает его в документ или сообщение, следуя указаниям в разделе, соответствующем вашему типу надстройки:

    - [Изменения, необходимые для надстройки Excel (JavaScript)](#changes-required-for-an-excel-add-in-javascript)
    - [Изменения, необходимые для надстройки Excel (TypeScript)](#changes-required-for-an-excel-add-in-typescript)
    - [Изменения, необходимые для надстройки Outlook (JavaScript)](#changes-required-for-an-outlook-add-in-javascript)
    - [Изменения, необходимые для надстройки Outlook (TypeScript)](#changes-required-for-an-outlook-add-in-typescript)
    - [Изменения, необходимые для надстройки PowerPoint (JavaScript)](#changes-required-for-a-powerpoint-add-in-javascript)
    - [Изменения, необходимые для надстройки PowerPoint (TypeScript)](#changes-required-for-a-powerpoint-add-in-typescript)
    - [Изменения, необходимые для надстройки Word (JavaScript)](#changes-required-for-a-word-add-in-javascript)
    - [Изменения, необходимые для надстройки Word (TypeScript)](#changes-required-for-a-word-add-in-typescript)

### <a name="changes-required-for-an-excel-add-in-javascript"></a>Изменения, необходимые для надстройки Excel (JavaScript)

Если ваша надстройка является надстройка Excel, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js:**

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-an-excel-add-in-typescript"></a>Изменения, необходимые для надстройки Excel (TypeScript)

Если ваша надстройка — это надстройка Excel, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией:

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-an-outlook-add-in-javascript"></a>Изменения, необходимые для надстройки Outlook (JavaScript)

Если ваша надстройка является надстройка Outlook, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js:**

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-an-outlook-add-in-typescript"></a>Изменения, необходимые для надстройки Outlook (TypeScript)

Если ваша надстройка — это надстройка Outlook, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией:

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-a-powerpoint-add-in-javascript"></a>Изменения, необходимые для надстройки PowerPoint (JavaScript)

Если ваша надстройка — это надстройка PowerPoint, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js:**

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-a-powerpoint-add-in-typescript"></a>Изменения, необходимые для надстройки PowerPoint (TypeScript)

Если ваша надстройка — это надстройка PowerPoint, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией:

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-a-word-add-in-javascript"></a>Изменения, необходимые для надстройки Word (JavaScript)

Если ваша надстройка — это надстройка Word, созданная с помощью JavaScript, внести следующие изменения в **./src/helpers/documentHelper.js:**

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

После внесения этих изменений переперейти к [](#try-it-out) разделу "Попробовать" этой статьи, чтобы проверить обновленную надстройки.

### <a name="changes-required-for-a-word-add-in-typescript"></a>Изменения, необходимые для надстройки Word (TypeScript)

Если ваша надстройка — это надстройка Word, созданная с помощью TypeScript, откройте **./src/taskpane/taskpane.ts,** найдите функцию и замените ее следующей `writeDataToOfficeDocument` функцией:

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

После внесения этих изменений переопробуйте обновленные надстройки в разделе "Попробуйте" этой статьи. [](#try-it-out)

## <a name="try-it-out"></a>Проверка

Если ваша надстройка является надстройка Excel, Word или PowerPoint, выполните действия из следующего раздела, чтобы опробовать ее. Если ваша надстройка является надстройка Outlook, выполните действия в разделе [Outlook.](#outlook)

### <a name="excel-word-and-powerpoint"></a>Excel, Word и PowerPoint

Выполните следующие действия, чтобы испытать надстройку Excel, Word или PowerPoint.

1. В корневой папке проекта запустите следующую команду для построения проекта, запуска локального веб-сервера и загрузки неопровержимой надстройки в ранее выбранном клиентом приложении Office.

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman.

    ```command&nbsp;line
    npm start
    ```

2. В клиентских приложениях Office, открываемых при запуске предыдущей команды (например, Excel, Word или PowerPoint), убедитесь, что вы вписались с пользователем, который является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которая использовалась для подключения к Azure при настройке [службы SSO](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа. 

3. В клиентском приложении Office выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки. На рисунке ниже показана эта кнопка в Excel. 

    ![Снимок экрана: выделенная кнопка надстройки на ленте Excel](../images/excel-quickstart-addin-3b.png)

4. В нижней части области задач выберите кнопку "Чтение **oneDrive** для бизнеса", чтобы инициировать процесс SSO.

5. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана: диалоговое окно, запрашивающее разрешения, с выделенной кнопкой "Принять"](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

6. Надстройка считывает данные из OneDrive для бизнеса во выгрузки пользователя и записывает в документ имена 10 самых верхних файлов и папок. На следующем рисунке показан пример имен файлов и папок, написанных на листах Excel.

    ![Screenshot showing OneDrive for Business information in Excel worksheet](../images/sso-onedrive-info-excel.png)

### <a name="outlook"></a>Outlook

Выполните следующие действия, чтобы испытать надстройку Outlook.

1. В корневой папке проекта запустите следующую команду для построения проекта, запуска локального веб-сервера и загрузки неогрузки надстройки. 

    > [!NOTE]
    > Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки. Если вам будет предложено установить сертификат после того, как вы запустите указанную ниже команду, примите предложение установить сертификат, предоставленный генератором Yeoman. Кроме того, вам может потребоваться запустить командную строку или терминал с правами администратора, чтобы внести изменения.

    ```command&nbsp;line
    npm start
    ```

2. Убедитесь, что вы вписались в Outlook с пользователем, который является членом той же организации Microsoft 365, что и учетная запись администратора Microsoft 365, которая использовалась для подключения к Azure при настройке [SSO](sso-quickstart.md#configure-sso) для приложения. Благодаря этому будут созданы соответствующие условия для успешного единого входа.

3. В Outlook создайте новое сообщение.

4. В окне создания сообщения нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: выделенная кнопка ленты надстройки в окне создания сообщения Outlook](../images/outlook-sso-ribbon-button.png)

5. В нижней части области задач выберите кнопку "Чтение **oneDrive** для бизнеса", чтобы инициировать процесс SSO.

6. Если открывается диалоговое окно, в котором запрашиваются разрешения от имени надстройки, это означает, что единый вход не поддерживается для вашего сценария и надстройка использует альтернативный метод проверки подлинности пользователя. Это может произойти, если администратор клиента не дал согласие на доступ надстройки к Microsoft Graph или если пользователь не вошел в Office с помощью действительной учетной записи Майкрософт или рабочей или учебной учетной записи Microsoft 365. Чтобы продолжить, нажмите кнопку **Принять** в диалоговом окне.

    ![Снимок экрана: диалоговое окно, запрашивающее разрешения, с выделенной кнопкой "Принять"](../images/sso-permissions-request.png)

    > [!NOTE]
    > После принятия пользователем запрос разрешений больше не выводится на экран.

7. Надстройка считывает данные из OneDrive для бизнеса во выгрузки и записывает имена 10 самых верхних файлов и папок в текст сообщения электронной почты.

    ![Screenshot showing OneDrive for Business information in Outlook compose message window](../images/sso-onedrive-info-outlook.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно настроили функции надстройки с поддержкой SSO, созданной с помощью генератора Yeoman в кратком запуске [SSO.](sso-quickstart.md) Дополнительные сведения об этапах настройки единого входа, которые генератор Yeoman выполняет автоматически, и коде, который упрощает процесс единого входа, см. в статье [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md).

## <a name="see-also"></a>См. также

- [Включение единого входа для надстроек Office](../develop/sso-in-office-add-ins.md)
- [Краткое руководство по единому входу (SSO)](sso-quickstart.md)
- [Создание надстройки Office на платформе Node.js с использованием единого входа](../develop/create-sso-office-add-ins-nodejs.md)
- [Устранение ошибок единого входа](../develop/troubleshoot-sso-in-office-add-ins.md)
