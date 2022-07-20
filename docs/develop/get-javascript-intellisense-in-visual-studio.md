---
title: Использование IntelliSense для JavaScript в Visual Studio
description: Узнайте, как использовать JSDoc для создания IntelliSense для переменных JavaScript, объектов, параметров и возвращаемых значений.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: deef6fe4356264534732e7f38a58a4079223686d
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889325"
---
# <a name="get-javascript-intellisense-in-visual-studio"></a>Использование IntelliSense для JavaScript в Visual Studio

При разработке надстроек Office с помощью Visual Studio 2019 и более поздних версий можно использовать JSDoc для включения IntelliSense для переменных JavaScript, объектов, параметров и возвращаемых значений. В этой статье представлен обзор JSDoc и способы его использования для создания IntellSense в Visual Studio. Дополнительные сведения см. в статье о поддержке [JavaScript IntelliSense](/visualstudio/ide/javascript-intellisense) и [JSDoc в JavaScript](https://github.com/Microsoft/TypeScript/wiki/JsDoc-support-in-JavaScript).

## <a name="officejs-type-definitions"></a>Определения типов Office.js

Вам необходимо предоставить Visual Studio определения типов Office.js. Для этого можно сделать следующее:

- Создать локальную копию файлов Office.js в папке вашего решения под названием `\Office\1\`. Эта локальная копия будет добавлена в шаблоны надстройки Office в Visual Studio во время создания проекта надстройки.
- Использовать интернет-версию Office.js, добавив файл tsconfig.json в корневой каталог проекта веб-приложения в решении надстройки. Этот файл должен иметь указанное ниже содержимое.

    ```json
        {
            "compilerOptions": {
                "allowJs": true,            // These settings apply to JavaScript files also.
                "noEmit":  true             // Do not compile the JS (or TS) files in this project.
            },
            "exclude": [
                "node_modules",             // Don't include any JavaScript found under "node_modules".
                "Scripts/Office/1"          // Suppress loading all the JavaScript files from the Office NuGet package.
            ],
            "typeAcquisition": {
                "enable": true,             // Enable automatic fetching of type definitions for detected JavaScript libraries.
                "include": [ "office-js" ]  // Ensure that the "Office-js" type definition is fetched.
            }
        }
    ```

## <a name="jsdoc-syntax"></a>Синтаксис JSDoc

Основной метод — добавить перед переменной (параметром и т. п.) комментарий с указанием типа данных. Это позволит IntelliSense в Visual Studio определять участников. Примеры.

### <a name="variable"></a>Переменная

```js
/** @type {Excel.Range} */
let subsetRange;
```

![Фрагмент IntelliSense для переменной subsetRange.](../images/intellisense-vs22-var.png)

### <a name="parameter"></a>Параметр

```js
/** @param {Word.ParagraphCollection} paragraphs */
function myFunc(paragraphs){

}
```

![Фрагмент intelliSense для параметра paras (параметр paragraphs в примере JavaScript).](../images/intellisense-vs17-param.png)

### <a name="return-value"></a>Возвращаемое значение

```js
/** @returns {Word.Range} */
function myFunc() {

}
```

![Фрагмент intelliSense для возвращаемого значения myFunc().](../images/intellisense-vs17-return.png)

### <a name="complex-types"></a>Сложные типы

```js
/** @typedef {{range: Word.Range, paragraphs: Word.ParagraphCollection}} MyType

/** @returns {MyType} */
function myFunc() {

}
```

![Например, IntelliSense для объявления сложного типа "let myVar;".](../images/intellisense-vs22-complex-type.png)

## <a name="see-also"></a>См. также

- [Разработка надстроек Office с помощью Visual Studio](develop-add-ins-visual-studio.md)
- [Отладка надстроек Office в Visual Studio](debug-office-add-ins-in-visual-studio.md)
