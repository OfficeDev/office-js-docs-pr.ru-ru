---
title: Модульное тестирование в надстройки Office
description: Узнайте, как выполнить модульное тестирование кода, который вызывает API JavaScript для Office.
ms.date: 02/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 21858a68734ca5d07621f3e9c88b147ebac7dde6
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958750"
---
# <a name="unit-testing-in-office-add-ins"></a>Модульное тестирование в надстройки Office

Модульные тесты проверяют функциональность надстройки без подключения к сети или службе, включая подключения к приложению Office. Модульное тестирование серверного кода и клиентского кода, не вызывающего [API JavaScript для Office](../develop/understanding-the-javascript-api-for-office.md), в надстройки Office совпадают с кодом в любом веб-приложении, поэтому для него не требуется специальная документация. Но код на стороне клиента, который вызывает API JavaScript для Office, сложно протестировать. Для решения этих проблем мы создали библиотеку для упрощения создания фиктивных объектов Office в модульных тестах [: Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock). Библиотека упрощает тестирование следующими способами:

- API JavaScript для Office должны инициализироваться в элементе управления webview в контексте приложения Office (Excel, Word и т. д.), поэтому их нельзя загрузить в процессе, в котором модульные тесты выполняются на компьютере разработки. Библиотеку Office-Addin-Mock можно импортировать в тестовые файлы, что позволяет макетировать API JavaScript для Office в node.js, в котором выполняются тесты.
- [API для конкретного](../develop/understanding-the-javascript-api-for-office.md#api-models) приложения имеют методы [загрузки](../develop/application-specific-api-model.md#load) и синхронизации, которые должны вызываться в определенном порядке относительно других функций и друг друга.[](../develop/application-specific-api-model.md#sync) Кроме того, `load` метод должен вызываться с определенными параметрами в зависимости от того, какие свойства объектов Office будут считываться кодом далее  в тестируемой функции. Однако платформы модульного тестирования по своей сути не имеют состояния, `load` `sync` поэтому они не могут записывать, были ли вызваны или какие параметры были переданы `load`. Фиктивные объекты, созданные с помощью библиотеки Office-Addin-Mock, имеют внутреннее состояние, которое отслеживает эти данные. Это позволяет фиктивным объектам эмулировать поведение ошибок фактических объектов Office. Например, если проверяемая `load`функция попытается прочитать свойство, которое не было передано в первый раз, тест вернет ошибку, аналогичную возвращаемой Office.

Библиотека не зависит от API JavaScript для Office и может использоваться с любой платформой модульного тестирования JavaScript, например:

- [Шутка](https://jestjs.io)
- [Мокко](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)

В примерах в этой статье используется платформа Jest. Примеры использования платформы Mocha можно использовать на [домашней странице Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#examples).

## <a name="prerequisites"></a>Предварительные условия

В этой статье предполагается, что вы знакомы с основными понятиями модульного тестирования и макетирования, включая создание и запуск тестовых файлов, а также что у вас есть опыт работы с платформой модульного тестирования.

> [!TIP]
> Если вы работаете с Visual Studio, рекомендуем ознакомиться со статьей модульного тестирования [JavaScript и TypeScript в Visual Studio](/visualstudio/javascript/unit-testing-javascript-with-visual-studio) , чтобы получить некоторые основные сведения о модульном тестировании JavaScript в Visual Studio, а затем вернуться к этой статье.

## <a name="install-the-tool"></a>Установка средства

Чтобы установить библиотеку, откройте командную строку, перейдите к корню проекта надстройки и введите следующую команду.

```command&nbsp;line
npm install office-addin-mock --save-dev
```

## <a name="basic-usage"></a>Базовое использование

1. Проект будет содержать один или несколько тестовых файлов. (См. инструкции для платформы тестирования и примеры файлов тестов в examples(#examples) ниже.) Импортируйте библиотеку с ключевым словом или ключевым словом в любой тестовый файл, содержащий тест функции, `require` `import` которая вызывает API JavaScript для Office, как показано в следующем примере.

   ```javascript
   const OfficeAddinMock = require("office-addin-mock");
   ```

1. Импортируйте модуль, содержащий функцию надстройки, которую требуется проверить с помощью ключевого `require` слова или ключевого `import` слова. Ниже приведен пример, в котором предполагается, что тестовый файл находится во вложенной папке папки с файлами кода надстройки.

   ```javascript
   const myOfficeAddinFeature = require("../my-office-add-in");
   ```

1. Создайте объект данных со свойствами и под свойствами, которые необходимо макетировать для тестирования функции. Ниже приведен пример объекта, который имитирует свойства Excel [Workbook.range.address](/javascript/api/excel/excel.range#excel-excel-range-address-member) и [метод Workbook.getSelectedRange](/javascript/api/excel/excel.workbook#excel-excel-workbook-getselectedrange-member(1)) . Это не окончательный макет объекта. Представьте, что это начальное значение, которое используется для `OfficeMockObject` создания конечного макета объекта.

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

1. Передайте объект данных в `OfficeMockObject` конструктор. Обратите внимание на следующие сведения о возвращаемом объекте `OfficeMockObject` .

   - Это упрощенный макет объекта [OfficeExtension.ClientRequestContext](/javascript/api/office/officeextension.clientrequestcontext) .
   - Макетный объект содержит все члены объекта данных, а также макетные реализации и `load` методы `sync` .
   - Макетный объект имитирует критическое поведение ошибки `ClientRequestContext` объекта. Например, если тестируемой API `sync`Office пытается прочитать свойство без предварительной загрузки свойства и вызова, то тест завершится ошибкой, аналогичной тому, что выдается в рабочей среде выполнения: "Ошибка, свойство не загружено".

   ```javascript
   const contextMock = new OfficeAddinMock.OfficeMockObject(mockData);
   ```

    > [!NOTE]
    > Полная справочная документация по типу `OfficeMockObject` находится на [сайте Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

1. В синтаксисе платформы тестирования добавьте тест функции. Используйте объект `OfficeMockObject` вместо макета объекта, в данном случае это `ClientRequestContext` объект. Далее продолжается пример в Jest. `getSelectedRangeAddress`В этом примере теста предполагается, что тестируемая функция надстройки вызывается, `ClientRequestContext` что она принимает объект в качестве параметра и предназначена для возврата адреса выбранного диапазона. Полный пример приведен [далее в этой статье](#mocking-a-clientrequestcontext-object).

   ```javascript
   test("getSelectedRangeAddress should return the address of the range", async function () {
     expect(await getSelectedRangeAddress(contextMock)).toBe("C2:G3");
   });
   ```

1. Запустите тест в соответствии с документацией по платформе тестирования и средствам разработки. Как правило, существует файл **package.json** со скриптом, который выполняет платформу тестирования. Например, если Jest является платформой, **package.json** будет содержать следующее:

   ```json
   "scripts": {
     "test": "jest",
     -- other scripts omitted --  
   }
   ```

   Чтобы запустить тест, введите следующее в командной строке в корневом каталоге проекта.

   ```command&nbsp;line
   npm test
   ```

## <a name="examples"></a>Примеры

В примерах в этом разделе используется Jest с параметрами по умолчанию. Эти параметры поддерживают модули CommonJS. Сведения о [том](https://jestjs.io/docs/getting-started) , как настроить Jest и node.js для поддержки модулей ECMAScript и TypeScript, см. в документации Jest. Чтобы выполнить любой из этих примеров, выполните следующие действия.

1. Создайте проект надстройки Office для соответствующего ведущего приложения Office (например, Excel или Word). Один из способов быстро сделать это — использовать генератор [Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).
1. В корне проекта установите [Jest](https://jestjs.io/docs/getting-started).
1. [Установите средство office-addin-mock](#install-the-tool).
1. Создайте файл точно так же, как первый файл в примере, и добавьте его в папку, содержащую другие исходные файлы проекта, которые часто называются `\src`.
1. Создайте вложенную папку в исходной папке и присвойте ему соответствующее имя, например `\tests`.
1. Создайте файл точно так же, как тестовый файл в примере, и добавьте его во вложенную папку.
1. Добавьте скрипт `test` в файл **package.json** , а затем запустите тест, как описано в разделе " [Базовое использование"](#basic-usage).

### <a name="mocking-the-office-common-apis"></a>Макетирование общих API Office

В этом примере предполагается, что надстройка Office для любого узла, поддерживающую общие [API Office](../develop/office-javascript-api-object-model.md) (например, Excel, PowerPoint или Word). Надстройка имеет одну из своих функций в файле с именем `my-common-api-add-in-feature.js`. Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!". к элементу, выбранному в документе; например, Диапазон в Word, ячейка в Excel или текстовое поле в PowerPoint.

```javascript
const myCommonAPIAddinFeature = {

    addHelloWorldText: async () => {
        const options = { coercionType: Office.CoercionType.Text };
        await Office.context.document.setSelectedDataAsync("Hello World!", options);
    }
}
  
module.exports = myCommonAPIAddinFeature;
```

Тестовый файл с именем `my-common-api-add-in-feature.test.js` находится во вложенной папке относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание `context`, что свойство верхнего уровня является объектом [Office.Context](/javascript/api/office/office.context) , поэтому объект, который имитируется, является родительским для этого свойства: [объект Office](/javascript/api/office) . Обратите внимание на следующие особенности этого кода:

- Конструктор `OfficeMockObject` *не добавляет все* классы перечисления Office `Office` в макетный объект, поэтому значение, на которое ссылаются в методе надстройки, `CoercionType.Text` должно быть явно добавлено в объект начального значения.
- Так как библиотека JavaScript для Office не загружена в процесс узла, объект, на который ссылаются в коде надстройки, `Office` должен быть объявлен и инициализирован.

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

### <a name="mocking-the-outlook-apis"></a>Макетирование API Outlook

Строго говоря, API Outlook являются частью модели общего API, но они имеют специальную архитектуру, созданную на основе объекта [Mailbox](/javascript/api/outlook/office.mailbox) , поэтому мы предоставили отдельный пример для Outlook. В этом примере предполагается, что Outlook имеет одну из функций в файле с именем `my-outlook-add-in-feature.js`. Ниже показано содержимое файла. Функция `addHelloWorldText` задает текст "Hello World!". к элементу, выбранному в настоящее время в окне создания сообщения.

```javascript
const myOutlookAddinFeature = {

    addHelloWorldText: async () => {
        Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
      }
}

module.exports = myOutlookAddinFeature;
```

Тестовый файл с именем `my-outlook-add-in-feature.test.js` находится во вложенной папке относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание `context`, что свойство верхнего уровня является объектом [Office.Context](/javascript/api/office/office.context) , поэтому объект, который имитируется, является родительским для этого свойства: [объект Office](/javascript/api/office) . Обратите внимание на следующие особенности этого кода:

- Свойство `host` макета объекта используется внутри библиотеки макета для идентификации приложения Office. Это обязательный параметр для Outlook. В настоящее время он не предназначен ни для каких других приложений Office.
- Так как библиотека JavaScript для Office не загружена в процесс узла, объект, на который ссылаются в коде надстройки, `Office` должен быть объявлен и инициализирован.

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

### <a name="mocking-the-office-application-specific-apis"></a>Макетирование API для конкретного приложения Office

При тестировании функций, использующих API для конкретного приложения, убедитесь, что вы макетируете правильный тип объекта. Возможны два варианта:

- Макет [OfficeExtension.ClientRequestObject](/javascript/api/office/officeextension.clientrequestcontext). Это можно сделать, если проверяемая функция соответствует обоим из следующих условий:

  - Он не *вызывает host.*`run` функция , например [Excel.run](/javascript/api/excel#Excel_run_batch_).
  - Он не ссылаются на любое другое прямое свойство или метод объекта *Host* .

- Макет объекта *Host* , например [Excel](/javascript/api/excel) или [Word](/javascript/api/word). Сделайте это, если предыдущий вариант невозможен.

Примеры обоих типов тестов приведены в подразделах ниже.

#### <a name="mocking-a-clientrequestcontext-object"></a>Макет объекта ClientRequestContext

В этом примере предполагается, что надстройка Excel имеет одну из функций в файле с именем `my-excel-add-in-feature.js`. Ниже показано содержимое файла. Обратите внимание, что `getSelectedRangeAddress` это вспомогательный метод, вызываемый в обратном вызове, который передается в `Excel.run`.

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

Тестовый файл с именем `my-excel-add-in-feature.test.js` находится во вложенной папке относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня имеет `workbook`значение, поэтому объект, который имитируется, `Excel.Workbook`является родительским для : объекта `ClientRequestContext` .

```javascript
const OfficeAddinMock = require("office-addin-mock");
const myExcelAddinFeature = require("../my-excel-add-in-feature");

// Create the seed mock object.
const mockData = {
    workbook: {
      range: {
        address: "C2:G3",
      },
      // Mock the Workbook.getSelectedRange method.
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

#### <a name="mocking-a-host-object"></a>Макетирование объекта узла

В этом примере предполагается, что надстройка Word имеет одну из своих функций в файле с именем `my-word-add-in-feature.js`. Ниже показано содержимое файла.

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

Тестовый файл с именем `my-word-add-in-feature.test.js` находится во вложенной папке относительно расположения файла кода надстройки. Ниже показано содержимое файла. Обратите внимание, что свойство верхнего уровня `context`является объектом `ClientRequestContext` , поэтому объект, который имитируется, является родительским для этого свойства: объекта `Word` . Обратите внимание на следующие особенности этого кода:

- Когда конструктор `OfficeMockObject` создает окончательный макетный объект, он гарантирует, `ClientRequestContext` что дочерний объект имеет и `sync` методы `load` .
- Конструктор `OfficeMockObject` не добавляет *функцию* в `run` `Word` макетный объект, поэтому ее необходимо явно добавить в объект начального значения.
- Конструктор `OfficeMockObject` *не добавляет все* классы перечисления Word `Word` в макетный объект, поэтому значение, на которое ссылаются в методе надстройки, `InsertLocation.end` должно быть явно добавлено в начальном объекте.
- Так как библиотека JavaScript для Office не загружена в процесс узла, объект, на который ссылаются в коде надстройки, `Word` должен быть объявлен и инициализирован.

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
  // Mock the Word.run function.
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
> Полная справочная документация по типу `OfficeMockObject` находится на [сайте Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock#reference).

## <a name="see-also"></a>Дополнительные ресурсы

- [Точка установки страницы npm в Office-Addin-Mock](https://www.npmjs.com/package/office-addin-mock) . 
- Репозиторий открытый код [Office-Addin-Mock](https://github.com/OfficeDev/Office-Addin-Scripts/tree/master/packages/office-addin-mock).
- [Шутка](https://jestjs.io)
- [Мокко](https://mochajs.org/)
- [Жасмин](https://jasmine.github.io/)
