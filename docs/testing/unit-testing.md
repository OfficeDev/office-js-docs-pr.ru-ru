---
title: Тестирование единиц в Office надстройки
description: Узнайте, как унифизировать тестовый код, который вызывает Office API JavaScript
ms.date: 11/30/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8824b8e759e3c1acecf30683f2b89bb41bd558f3
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242042"
---
# <a name="unit-testing-in-office-add-ins"></a>Тестирование единиц в Office надстройки

Unit tests check your add-in's functionality without requiring network or service connections, including connections to the Office application. Код, тестируемый на стороне сервера, и клиентский код, который не называет API [javaScript](../develop/understanding-the-javascript-api-for-office.md)Office, в Office надстройки такие же, как и в любом веб-приложении, поэтому для этого не требуется специальная документация.  Но клиентский код, который вызывает Office API JavaScript, тестировать сложно. Чтобы решить эти проблемы, мы создали библиотеку, чтобы упростить создание объектов макета Office в unit tests: [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). Библиотека упрощает тестирование следующими способами:

- API Office JavaScript должны инициализироваться в области управления веб-просмотром в контексте приложения Office (Excel, Word и т.д.), поэтому они не могут быть загружены в процессе, в котором блок-тесты запускаются на компьютере разработки. Библиотека Office-Addin-Mock может быть импортирована в тестовые файлы, что позволяет высмеять Office API JavaScript внутри node.js, в котором тесты запускаются.
- API, [определенные приложениям,](../develop/understanding-the-javascript-api-for-office.md#api-models) имеют методы загрузки и синхронизации, которые должны быть вызваны в определенном порядке относительно других функций и друг к другу. [](../develop/application-specific-api-model.md#load) [](../develop/application-specific-api-model.md#sync) Кроме того, метод должен быть вызван с определенными параметрами в зависимости от того, какие свойства Office объектов будут считыты кодом позже в проверяемой `load` функции.  Но фреймворки для тестирования единиц по своей сути не имеют состояния, поэтому они не могут вести учет того, были ли вызваны или какие параметры переданы `load` `sync` `load` . Объекты макета, которые вы создаете с библиотекой Office-Addin-Mock, имеют внутреннее состояние, которое отслеживает эти вещи. Это позволяет макету объектов эмулировать поведение ошибки фактических Office объектов. Например, если проверяемая функция пытается прочитать свойство, которое не было впервые передано, то тест возвращает ошибку, аналогичную Office `load` возвращается.

Библиотека не зависит от API Office JavaScript, и ее можно использовать с любой платформой тестирования подразделений JavaScript, например:

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)

В примерах этой статьи используется фреймворк Jest. Примеры использования фреймворка Mocha на домашней [странице Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Предварительные требования

В этой статье предполагается, что вы знакомы с основными понятиями тестирования и макетов единиц, включая создание и запуск тестовых файлов, и что у вас есть некоторый опыт работы с инфраструктурой тестирования единицы.

> [!TIP]
> Если вы работаете с Visual Studio, рекомендуем прочитать статью Unit testing JavaScript и TypeScript в Visual Studio для некоторых базовых сведений о тестировании подразделений [JavaScript](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) в Visual Studio, а затем вернуться к этой статье.

## <a name="install-the-tool"></a>Установка средства

Чтобы установить библиотеку, откройте командную подсказку, перейдите к корню проекта надстройки и введите следующую команду.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Базовое использование

1. В проекте будет один или несколько тестовых файлов. (См. инструкции для тестовой базы и примеры тестовых файлов в примерах (#examples) ниже.) Импортировать библиотеку с помощью ключевого слова или ключевого слова в любой тестовый файл с тестом функции, которая вызывает Office API JavaScript, как показано в следующем `require` `import` примере.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Импортировать модуль, содержащий функцию надстройки, которую необходимо протестировать с помощью `require` ключевого слова или `import` ключевого слова. Ниже приводится пример, который предполагает, что тестовый файл находится в подмостках папки с файлами кода надстройки.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Создайте объект данных с свойствами и свойствами, которые необходимо инсценировать для проверки функции. Ниже приводится пример объекта, высмеять Excel [Workbook.range.address](/javascript/api/excel/excel.range#address) и метода [Workbook.getSelectedRange.](/javascript/api/excel/excel.workbook#getSelectedRange__) Это не последний объект макета. Подумайте об этом как о объекте seed, который используется для `OfficeMockObject` создания конечного объекта макета.

   ```javascript
   const mockData = {
     workbook: {
       range: {
         address: "C2:G3",
       },
       getSelectedRange: function () {
         return this.range;
       },
     },
   };
   ```

1. Передай объект данных `OfficeMockObject` конструктору. Обратите внимание на следующее о возвращенных `OfficeMockObject` объектах.

   - Это упрощенный макет объекта [OfficeExtension.ClientRequestContext.](/javascript/api/office/officeextension.clientrequestcontext)
   - Макет объекта имеет все члены объекта данных, а также макет реализаций `load` и `sync` методов.
   - Макет объекта будет имитировать критическое поведение ошибки `ClientRequestContext` объекта. Например, если API Office, который вы тестируете, пытается прочитать свойство без первой загрузки свойства и вызова, то тест будет сбой с ошибкой, аналогичной тому, что будет брошено в производственное время: "Ошибка, свойство не `sync` загружено".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > Полная справочная документация по типу `OfficeMockObject` находится [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. В синтаксис тестовой базы добавьте тест функции. Используйте объект `OfficeMockObject` на месте объекта, который он макет, в этом случае `ClientRequestContext` объект. Далее приводится пример в Jest. В этом примере тест предполагает, что проверяемая функция надстройки называется, что она принимает объект в качестве параметра и что она предназначена для возврата адреса выбранного в настоящее время `getSelectedRangeAddress` `ClientRequestContext` диапазона. Полный пример [далее в этой статье](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Запустите тест в соответствии с документацией по тестовой инфраструктуре и средствам разработки. Как правило, существует **файл package.json** со сценарием, который выполняет тестовую базу. Например, если Jest является фреймворком, **package.json будет** содержать следующее:

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   Чтобы выполнить тест, введите следующее в командной подсказке в корне проекта.

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>Примеры

В примерах этого раздела используется Jest с его настройками по умолчанию. Эти параметры поддерживают модули CommonJS. См. [документацию Jest](https://jestjs.io/docs/getting-started) о настройке Jest и node.js для поддержки модулей ECMAScript и поддержки TypeScript. Чтобы выполнить любой из этих примеров, необходимо выполнить следующие действия.

1. Создайте Office надстройки для соответствующего Office хост-приложения (например, Excel или Word). Один из способов быстро сделать это — использовать [средство Yo Office.](https://github.com/OfficeDev/generator-office)
1. В корне проекта установите [Jest](https://jestjs.io/docs/getting-started).
1. [Установите средство office-addin-mock.](#install-the-tool)
1. Создайте файл точно так же, как первый файл в примере, и добавьте его в папку, которая содержит другие исходные файлы проекта, часто называемые `\src` .
1. Создайте подмостки в папку исходных файлов и назови ей соответствующее имя, например `\tests` .
1. Создайте файл точно так же, как тестовый файл в примере, и добавьте его в подмостки.
1. Добавьте скрипт `test` в **файл package.json** и запустите тест, как описано в [базовом использовании.](#basic-usage)

### <a name="mocking-the-office-common-apis"></a>Макет Office API

В этом примере предполагается Office надстройка для любого хоста, поддерживающего Office [API](../develop/office-javascript-api-object-model.md) (например, Excel, PowerPoint или Word). Надстройка имеет одну из своих функций в файле с именем `my-common-api-add-in-feature.js` . Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!" все, что в настоящее время выбрано в документе; например; диапазон в Word или ячейка в Excel или текстовое поле в PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

Тестовый файл, названный в подмостке, относительно расположения файла `my-common-api-add-in-feature.test.js` кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `context` — это [Office. Объект Context,](/javascript/api/office/office.context) поэтому объект, на который ведется насмешка, является родителем этого [свойства: Office](/javascript/api/office) объектом. Обратите внимание на следующие особенности этого кода:

- Конструктор не добавляет все классы Office в макетный объект, поэтому значение, которое ссылается в методе надстройки, должно быть явно добавлено в объект `OfficeMockObject`  `Office` `CoercionType.Text` семени.
- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, должен быть объявлен и `Office` инициализирован.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myCommonAPIAddinFeature = require("../my-common-api-add-in-feature");

// Create the seed mock object.
const mockData = {
    context: {
      document: {
        setSelectedDataAsync: function (data, options) {
          this.data = data;
          this.options = options;
        },
      },
    },
    // Mock the Office.CoercionType enum.
    CoercionType: {
      Text: {},
    },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in document should be set to 'Hello World'", async function () {
    await myCommonAPIAddinFeature.addHelloWorldText();
    expect(officeMock.context.document.data).toBe("Hello World!");
});
```

### <a name="mocking-the-outlook-apis"></a>Макет Outlook API

Хотя строго говоря, Outlook API являются частью общей модели API, они имеют специальную архитектуру, которая построена вокруг объекта [почтовых](/javascript/api/outlook/office.mailbox) ящиков, поэтому мы предоставили отдельный пример для Outlook. В этом примере предполагается Outlook, который имеет одну из своих функций в файле с именем `my-outlook-add-in-feature.js` . Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!" на все, что в настоящее время выбрано в окне составить сообщение.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

Тестовый файл, названный в подмостке, относительно расположения файла `my-outlook-add-in-feature.test.js` кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `context` — это [Office. Объект Context,](/javascript/api/office/office.context) поэтому объект, на который ведется насмешка, является родителем этого [свойства: Office](/javascript/api/office) объектом. Обратите внимание на следующие особенности этого кода:

- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, должен быть объявлен и `Office` инициализирован.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    mailbox: {
      item: {
          setSelectedDataAsync: function (data) {
          this.data = data;
        },
      },
    },
  },
};
  
// Create the final mock object from the seed object.
const officeMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Create the Office object that is called in the addHelloWorldText function.
global.Office = officeMock;

/* Code that calls the test framework goes below this line. */

// Jest test
test("Text of selection in message should be set to 'Hello World'", async function () {
    await myOutlookAddinFeature.addHelloWorldText();
    expect(officeMock.context.mailbox.item.data).toBe("Hello World!");
});
```

### <a name="mocking-the-office-application-specific-apis"></a>Макет API Office приложений

При тестировании функций, которые используют API, определенные приложениям, убедитесь, что вы издевались над нужным типом объекта. Возможны два варианта:

- Макет [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Делайте это, когда проверяемая функция соответствует следующим условиям:

  - Он не называет *хост.*`run` метод, например [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Он не ссылается на любое другое прямое свойство или метод *объекта Host.*

- Макет объекта *Host,* например [Excel](/javascript/api/excel) [Word.](/javascript/api/word) Сделайте это, если предыдущий вариант невозможен.

Примеры тестов обоих типов приведены в подсекциях ниже.

#### <a name="mocking-a-clientrequestcontext-object"></a>Макет объекта ClientRequestContext

В этом примере предполагается Excel надстройка, которая имеет одну из своих функций в файле с именем `my-excel-add-in-feature.js` . Ниже показано содержимое файла. Обратите `getSelectedRangeAddress` внимание, что это метод помощника, вызываемого внутри отозвался, который передается `Excel.run` .

```javascript
const myExcelAddinFeature = {
    
    getSelectedRangeAddress: async (context) => {
        const range = context.workbook.getSelectedRange();      
        range.load("address");

        await context.sync();
      
        return range.address;
    }
}

module.exports = myExcelAddinFeature;
```

Тестовый файл, названный в подмостке, относительно расположения файла `my-excel-add-in-feature.test.js` кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня является таким образом, что объект, на который ведется насмешка, является родителем `workbook` `Excel.Workbook` объекта: `ClientRequestContext` объекта.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectRange method.
      getSelectedRange: function () {
        return this.range;
      },
    },
};

// Create the final mock object from the seed object.
const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);

/* Code that calls the test framework goes below this line. */

// Jest test
test("getSelectedRangeAddress should return address of selected range", async function () {
  expect(await myOfficeAddinFeature.getSelectedRangeAddress(contextMock)).toBe("C2:G3");
});
```

#### <a name="mocking-a-host-object"></a>Макет объекта-хоста

В этом примере предполагается надстройка Word, которая имеет одну из своих функций в файле с именем `my-word-add-in-feature.js` . Ниже показано содержимое файла.

```javascript
const myWordAddinFeature = {

  insertBlueParagraph: async () => {
    return Word.run(async (context) => {
      // Insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
  
      // Change the font color to blue.
      paragraph.font.color = "blue";
  
      await context.sync();
    });
  }
}

module.exports = myWordAddinFeature;
```

Тестовый файл, названный в подмостке, относительно расположения файла `my-word-add-in-feature.test.js` кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня — это объект, поэтому объект, на который ведется насмешка, является родителем этого `context` `ClientRequestContext` `Word` свойства: объектом. Обратите внимание на следующие особенности этого кода:

- Когда конструктор создает конечный макет объекта, он гарантирует, что у детского объекта `OfficeMockObject` `ClientRequestContext` есть и `sync` `load` методы.
- Конструктор `OfficeMockObject` не *добавляет метод* к объекту макета, поэтому он должен быть явно добавлен `run` в объект `Word` семени.
- Конструктор не добавляет все классы word enum к объекту макета, поэтому значение, которое ссылается в методе надстройки, должно быть явно добавлено в объект `OfficeMockObject`  `Word` `InsertLocation.end` семени.
- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, должен быть объявлен и `Word` инициализирован.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myWordAddinFeature = require("../my-word-add-in-feature");

// Create the seed mock object.
const mockData = {
  context: {
    document: {
      body: {
        paragraph: {
          font: {},
        },
        // Mock the Body.insertParagraph method.
        insertParagraph: function (paragraphText, insertLocation) {
          this.paragraph.text = paragraphText;
          this.paragraph.insertLocation = insertLocation;
          return this.paragraph;
        },
      },
    },
  },
  // Mock the Word.InsertLocation enum.
  InsertLocation: {
    end: "end",
  },
  // Mock the Word.run method.
  run: async function(callback) {
    await callback(this.context);
  },
};

// Create the final mock object from the seed object.
const wordMock = new OfficeAddinMock.OfficeMockObject(mockData);

// Define and initialize the Word object that is called in the insertBlueParagraph function.
global.Word = wordMock;

/* Code that calls the test framework goes below this line. */

// Jest test set
describe("Insert blue paragraph at end tests", () => {

  test("color of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();  
    expect(wordMock.context.document.body.paragraph.font.color).toBe("blue");
  });

  test("text of paragraph", async function () {
    await myWordAddinFeature.insertBlueParagraph();
    expect(wordMock.context.document.body.paragraph.text).toBe("Hello World");
  });
})
```

## <a name="adding-mock-objects-properties-and-methods-dynamically-when-testing"></a>Динамическое добавление макетных объектов, свойств и методов при тестировании

В некоторых сценариях эффективное тестирование требует создания или изменения макетных объектов во время работы; то есть, пока тесты запущены. Ниже представлены примеры.

- Проверяемая функция ведет себя по-разному, когда ее называют второй раз. Сначала необходимо протестировать функцию с помощью одного макетного объекта, а затем изменить этот объект макета и снова протестировать функцию с измененным объектом макета.
- Необходимо протестировать функцию с несколькими похожими, но не идентичными макетами объектов. Например, необходимо протестировать функцию с макетным объектом с свойством цвета, а затем снова протестировать функцию с макетным объектом с текстовым свойством, но в противном случае идентичным исходному объекту макета.

В `OfficeMockObject` этих сценариях можно использовать три метода.

- `OfficeMockObject.setMock` добавляет свойство и значение в `OfficeMockObject` объект. В следующем примере `address` добавляется свойство.

    ```javascript
    rangeMock.setMock("address", "G6:K9");
    ```

- `OfficeMockObject.addMockFunction` добавляет в объект макетную `OfficeMockObject` функцию, как показано в следующем примере.

    ```javascript
    workbookMock.addMockFunction("getSelectedRange", function () { 
      const range = {
        address: "B2:G5",
      };
      return range;
    });
    ```

    > [!NOTE]
    > Параметр функции необязателен. Если ее нет, создается пустая функция.

- `OfficeMockObject.addMock` добавляет новый `OfficeMockObject` объект в качестве свойства к существующему и дает ему имя. Он будет иметь минимальные члены, которые все `OfficeMockObject` имеют, такие как `load` и `sync` . Дополнительные члены могут быть добавлены с помощью `setMock` методов и `addMockFunction` методов. Ниже приводится пример, который добавляет объект макета в качестве свойства в `Excel.WorkbookProtection` `protection` макетную книгу. Затем добавляет свойство `protected` к новому объекту макета.

    ```javascript
    workbookMock.addMock("protection");
    workbookMock.protection.setMock("protected", true);
    ```

> [!NOTE]
> Полная справочная документация по типу `OfficeMockObject` находится [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>См. также

- [Office-Addin-Mock пункт установки страницы npm.](https://www.npmjs.com/package/office-addin-mock) 
- Репо с открытым исходным кодом [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)
