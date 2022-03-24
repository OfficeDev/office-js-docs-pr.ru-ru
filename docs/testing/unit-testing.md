---
title: Тестирование единиц в Office надстройки
description: Узнайте, как унифьмировать тестовый код, который вызывает Office API JavaScript.
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 6f6b0483b23c3f7199a8bd308bf8a4118402ee08
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746129"
---
# <a name="unit-testing-in-office-add-ins"></a>Тестирование единиц в Office надстройки

Unit tests check your add-in's functionality without requiring network or service connections, including connections to the Office application. Код, тестируемый на стороне сервера, и клиентский  код, который не называет API [javaScript](../develop/understanding-the-javascript-api-for-office.md) Office, в Office надстройки такие же, как и в любом веб-приложении, поэтому для этого не требуется специальная документация. Но клиентский код, который вызывает Office API JavaScript, тестировать сложно. Чтобы решить эти проблемы, мы создали библиотеку, чтобы упростить создание макетных объектов Office в unit tests: [Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). Библиотека упрощает тестирование следующими способами:

- API Office JavaScript должны инициализироваться в области управления веб-просмотром в контексте приложения Office (Excel, Word и т.д.), поэтому они не могут быть загружены в процессе, в котором на компьютере разработки запускаются тесты единиц. Библиотека Office-Addin-Mock может быть импортирована в тестовые файлы, что позволяет высмеять Office API JavaScript внутри node.js процесса, в котором тесты запускаются.
- API[, определенные приложениям](../develop/understanding-the-javascript-api-for-office.md#api-models), [](../develop/application-specific-api-model.md#load) имеют методы [](../develop/application-specific-api-model.md#sync) загрузки и синхронизации, которые должны быть вызваны в определенном порядке относительно других функций и друг к другу. Кроме того, `load` метод должен быть вызван с определенными параметрами в зависимости от того, какие свойства Office объектов будут считыты кодом позже в проверяемой функции. Но фреймворки для тестирования единиц по своей сути не имеют состояния, `load` `sync` `load`поэтому они не могут вести учет того, были ли вызваны или какие параметры переданы . Объекты макета, которые вы создаете с библиотекой Office-Addin-Mock, имеют внутреннее состояние, которое отслеживает эти вещи. Это позволяет макету объектов подражать поведению ошибок фактических Office объектов. Например, если `load`проверяемая функция пытается прочитать свойство, которое не было впервые передано, то тест возвращает ошибку, аналогичную тому, Office возвращается.

Библиотека не зависит от API Office JavaScript и может использоваться с любыми платформами тестирования подразделений JavaScript, такими как:

- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)

В примерах этой статьи используется фреймворк Jest. Примеры использования фреймворка Mocha на домашней [странице Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Предварительные требования

В этой статье предполагается, что вы знакомы с основными понятиями тестирования и макетов единиц, включая создание и запуск тестовых файлов, и что у вас есть некоторый опыт работы с инфраструктурой тестирования единицы.

> [!TIP]
> Если вы работаете с Visual Studio, рекомендуем прочитать статью Unit testing [JavaScript и TypeScript в Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) для некоторых базовых сведений о тестировании подразделений JavaScript в Visual Studio, а затем вернуться к этой статье.

## <a name="install-the-tool"></a>Установка средства

Чтобы установить библиотеку, откройте командную подсказку, перейдите к корню проекта надстройки и введите следующую команду.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Базовое использование

1. В проекте будет один или несколько тестовых файлов. (См. инструкции для тестовой базы и примеры тестовых файлов в примерах (#examples) ниже.) Импортировать библиотеку `require` `import` с помощью ключевого слова или ключевого слова в любой тестовый файл с тестом функции, которая вызывает Office API JavaScript, как показано в следующем примере.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Импортировать модуль, содержащий функцию надстройки, которую необходимо протестировать `require` с помощью ключевого слова или ключевого слова `import` . Ниже приводится пример, который предполагает, что тестовый файл находится в подмостках папки с файлами кода надстройки.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Создайте объект данных с свойствами и свойствами, которые необходимо инсценировать для проверки функции. Ниже приводится пример объекта, высмеять Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) и метод [Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)). Это не последний объект макета. Подумайте об этом как о объекте seed, который используется для `OfficeMockObject` создания конечного объекта макета.

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

1. Передай объект данных конструктору `OfficeMockObject` . Обратите внимание на следующее о возвращенных объектах `OfficeMockObject` .

   - Это упрощенный макет объекта [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) .
   - Макет объекта имеет все члены объекта данных, а также макет реализаций и `load` методов `sync` .
   - Макет объекта будет имитировать критическое поведение ошибки `ClientRequestContext` объекта. Например, если тестируемая Office API `sync`пытается прочитать свойство без первой загрузки свойства и вызова, то тест не пройдет с ошибкой, похожей на то, что было бы брошено в производственное время: "Ошибка, свойство не загружено".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > Полная справочная документация `OfficeMockObject` для этого [типа находится Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. В синтаксис тестовой базы добавьте тест функции. Используйте объект `OfficeMockObject` на месте объекта, который он макет, в этом случае `ClientRequestContext` объект. Далее приводится пример в Jest. В этом `getSelectedRangeAddress`примере тест предполагает, что проверяемая функция надстройки называется, `ClientRequestContext` что она принимает объект в качестве параметра и что она предназначена для возврата адреса выбранного в настоящее время диапазона. Полный пример [далее в этой статье](#mocking-a-clientrequestcontext-object).

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

1. Создайте Office надстройки для соответствующего Office хост-приложения (например, Excel или Word). Один из способов сделать это быстро — использовать генератор [Yeoman для Office надстройки](../develop/yeoman-generator-overview.md).
1. В корне проекта установите [Jest](https://jestjs.io/docs/getting-started).
1. [Установите средство office-addin-mock](#install-the-tool).
1. Создайте файл точно так же, как первый файл в примере, и добавьте его в папку, которая содержит другие исходные файлы проекта, часто называемые `\src`.
1. Создайте подмостки в папку исходных файлов и назови ей соответствующее имя, например `\tests`.
1. Создайте файл точно так же, как тестовый файл в примере, и добавьте его в подмостки.
1. Добавьте скрипт `test` в **файл package.json** и запустите тест, как описано в [базовом использовании](#basic-usage).

### <a name="mocking-the-office-common-apis"></a>Макет общих Office API

В этом примере предполагается Office надстройка для любого хоста, поддерживающего Office [API](../develop/office-javascript-api-object-model.md) (например, Excel, PowerPoint или Word). Надстройка имеет одну из своих функций в файле с именем `my-common-api-add-in-feature.js`. Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!" все, что в настоящее время выбрано в документе; например; диапазон в Word или ячейка в Excel или текстовое поле в PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

Тестовый файл, названный `my-common-api-add-in-feature.test.js` в подмостке, относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `context`— это [Office. Объект Context](/javascript/api/office/office.context), поэтому объект, на который ведется насмешка, является родителем этого [свойства: Office](/javascript/api/office) объектом. Обратите внимание на следующие особенности этого кода:

- Конструктор `OfficeMockObject` не добавляет все  классы Office `Office` в макетный объект, `CoercionType.Text` поэтому значение, которое ссылается в методе надстройки, должно быть явно добавлено в объект семени.
- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, `Office` должен быть объявлен и инициализирован.

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

Хотя строго говоря, Outlook API являются частью общей модели API, они имеют специальную архитектуру, которая строится вокруг объекта [почтовых](/javascript/api/outlook/office.mailbox) ящиков, поэтому мы предоставили отдельный пример для Outlook. В этом примере предполагается Outlook, который имеет одну из своих функций в файле с именем `my-outlook-add-in-feature.js`. Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!" на все, что в настоящее время выбрано в окне составить сообщение.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

Тестовый файл, названный `my-outlook-add-in-feature.test.js` в подмостке, относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `context`— это [Office. Объект Context](/javascript/api/office/office.context), поэтому объект, на который ведется насмешка, является родителем этого [свойства: Office](/javascript/api/office) объектом. Обратите внимание на следующие особенности этого кода:

- Свойство `host` на макете используется внутренне макетной библиотекой для идентификации Office приложения. Это обязательно для Outlook. В настоящее время он не предназначен для любого другого Office приложения.
- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, `Office` должен быть объявлен и инициализирован.

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myOutlookAddinFeature = require("../my-outlook-add-in-feature");

// Create the seed mock object.
const mockData = {
  // Identify the host to the mock library (required for Outlook).
  host: "outlook",
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

  - Он не вызываем *хост*.`run` метод, например [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Он не ссылается на любое другое прямое свойство или метод *объекта Host* .

- Макет объекта *Host*, например [Excel](/javascript/api/excel) [word](/javascript/api/word). Сделайте это, если предыдущий вариант невозможен.

Примеры тестов обоих типов приведены в подсекциях ниже.

#### <a name="mocking-a-clientrequestcontext-object"></a>Макет объекта ClientRequestContext

В этом примере предполагается Excel надстройка, которая имеет одну из своих функций в файле с именем `my-excel-add-in-feature.js`. Ниже показано содержимое файла. Обратите внимание, что `getSelectedRangeAddress` это метод помощника, вызываемого внутри отозвался, который передается .`Excel.run`

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

Тестовый файл, названный `my-excel-add-in-feature.test.js` в подмостке, относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `workbook`является таким `Excel.Workbook`образом, что объект, на который ведется насмешка, является родителем объекта: `ClientRequestContext` объекта.

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

В этом примере предполагается надстройка Word, которая имеет одну из своих функций в файле с именем `my-word-add-in-feature.js`. Ниже показано содержимое файла.

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

Тестовый файл, названный `my-word-add-in-feature.test.js` в подмостке, относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что `context`свойство верхнего уровня — `ClientRequestContext` это объект, поэтому объект, на который ведется насмешка, является родителем этого свойства: объектом `Word` . Обратите внимание на следующие особенности этого кода:

- Когда конструктор `OfficeMockObject` создает конечный макет объекта, он гарантирует `ClientRequestContext` , что у детского объекта есть и `sync` методы `load` .
- Конструктор `OfficeMockObject` не *добавляет метод* к объекту `run` `Word` макета, поэтому он должен быть явно добавлен в объект семени.
- Конструктор `OfficeMockObject` не добавляет все  классы word enum `Word` к объекту макета, поэтому значение, которое ссылается в методе надстройки, `InsertLocation.end` должно быть явно добавлено в объект семени.
- Так как Office JavaScript не загружается в процесс узла, объект, который ссылается в коде надстройки, `Word` должен быть объявлен и инициализирован.

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

> [!NOTE]
> Полная справочная документация `OfficeMockObject` для этого [типа находится Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>См. также

- [Office-Addin-Mock пункт установки страницы npm](https://www.npmjs.com/package/office-addin-mock). 
- Репо с открытым исходным [кодом Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Jest](https://jestjs.io)
- [Mocha](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)
