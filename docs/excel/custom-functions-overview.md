---
ms.date: 10/14/2020
description: Создайте пользовательскую функцию Excel для надстройки Office
title: Создание пользовательских функций в Excel
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 466050a5323f0f02fb886c763f5a2a594a9e2233
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48741115"
---
# <a name="create-custom-functions-in-excel"></a>Создание пользовательских функций в Excel

Пользовательские функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

Ниже на анимированном изображении показано, как рабочая книга вызывает функцию, созданную вами с помощью JavaScript или Typescript. В этом примере пользовательская функция `=MYFUNCTION.SPHEREVOLUME` рассчитывает объем сферы.

<img alt="animated image showing an end user inserting the MYFUNCTION.SPHEREVOLUME custom function into a cell of an Excel worksheet" src="../images/SphereVolumeNew.gif" />

Приведенный ниже код определяет пользовательскую функцию `=MYFUNCTION.SPHEREVOLUME`.

```js
/**
 * Returns the volume of a sphere.
 * @customfunction
 * @param {number} radius
 */
function sphereVolume(radius) {
  return Math.pow(radius, 3) * 4 * Math.PI / 3;
}
```

> [!NOTE]
> В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.

## <a name="how-a-custom-function-is-defined-in-code"></a>Как определена пользовательская функция в коде

Если использовать [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, он создаст файлы, управляющие вашими функциями и областью задач. Мы сосредоточимся на файлах, которые важны для пользовательских функций:

| Файл | Формат файла | Описание |
|------|-------------|-------------|
| **./src/functions/functions.js**<br/>или<br/>**./src/functions/functions.ts** | JavaScript<br/>или<br/>TypeScript | Содержит код, который определяет пользовательские функции. |
| **./src/functions/functions.html** | HTML | Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции. |
| **./manifest.xml** | XML | Указывает расположение нескольких файлов, которые используются пользовательскими функциями, например JavaScript, JSON и HTML-файлов. А также среду выполнения, которую должны использовать пользовательские функции, расположение файлов области задач и командных файлов. |

### <a name="script-file"></a>Файл скрипта

Файл скрипта (**./src/functions/functions.js** или **./src/functions/functions.ts**) содержит код, определяющий пользовательские функции, и комментарии, определяющие функцию.

Приведенный ниже код определяет пользовательскую функцию `add`. Примечания кода используются для создания файла метаданных JSON с описанием пользовательской функции для Excel. Обязательный комментарий `@customfunction` объявлен первым, чтобы указать, что это пользовательская функция. Затем объявляются еще два параметра: `first` и `second`, за которыми следуют их свойства `description`. Наконец, дается описание `returns`. Дополнительные сведения о том, какие комментарии являются обязательными для вашей пользовательской функции, см. в статье [Создание метаданных JSON для пользовательских функций](custom-functions-json-autogeneration.md).

```js
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number.
 * @param second Second number.
 * @returns The sum of the two numbers.
 */

function add(first, second){
  return first + second;
}
```

### <a name="manifest-file"></a>Файл манифеста

Файл манифеста XML для надстройки, определяющий пользовательские функции (**./manifest.xml** в проекте, созданном генератором Yo Office) выполняет следующее:

- Определяет пространство имен для пользовательских функций. Пространство имен добавляется к пользовательским функциям, чтобы клиенты могли определить ваши функции в рамках надстройки.
- Использует уникальные для манифеста пользовательских функций элементы `<ExtensionPoint>` и `<Resources>`. Эти элементы содержат сведения о расположении JavaScript, JSON и HTML-файлов.
- Указывает, какую среду выполнения использовать для пользовательской функции. Рекомендуется всегда использовать общую среду выполнения, если нет особой потребности в использовании другой среды, так как общая позволяет делиться данными между функциями и областью задач. Обратите внимание, что использование общей среды выполнения означает, что ваша надстройка будет использовать Internet Explorer 11, а не Microsoft Edge.

Если для создания файлов используется генератор Yo Office, рекомендуется настроить манифест для использования общей среды выполнения, так как это не настроено по умолчанию для этих файлов. Чтобы изменить манифест, следуйте инструкциям в статье [Настройка надстройки Excel для использования общей среды выполнения JavaScript](configure-your-add-in-to-use-a-shared-runtime.md).

Чтобы просмотреть полный рабочий манифест из примера надстройки, см. [этот репозиторий GitHub](https://github.com/OfficeDev/PnP-OfficeAddins/blob/master/Samples/excel-shared-runtime-global-state/manifest.xml).

[!include[manifest guidance](../includes/manifest-guidance.md)]

## <a name="coauthoring"></a>Совместное редактирование

Excel в Интернете и Windows, подключенные к подписке Microsoft 365, позволяют использовать совместное редактирование в Excel. Если в книге используется пользовательская функция, вашему коллеге по совместному редактированию будет предложено загрузить надстройку пользовательской функции. Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.

Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).

## <a name="known-issues"></a>Известные проблемы

С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues).

## <a name="next-steps"></a>Дальнейшие действия

Хотите попробовать пользовательские функции? Ознакомьтесь с простым [кратким руководством по началу работы с пользовательскими функциями](../quickstarts/excel-custom-functions-quickstart.md) или с более глубоким [руководством по пользовательским функциям](../tutorials/excel-tutorial-create-custom-functions.md), если вы этого еще не сделали.

Еще одно простое средство ознакомления с пользовательскими функциями — [Script Lab](https://appsource.microsoft.com/product/office/WA104380862?src=office&corrid=1ada79ac-6392-438d-bb16-fce6994a2a7e&omexanonuid=f7b03101-ec22-4270-a274-bcf16c762039&referralurl=https%3a%2f%2fgithub.com%2fofficedev%2fscript-lab), надстройка, в которой можно экспериментировать с пользовательскими функциями прямо в Excel. Вы можете попробовать создать собственные пользовательские функции или поиграть с готовыми примерами.

## <a name="see-also"></a>См. также 
* [Сведения о программе для разработчиков Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
* [Требования к настраиваемым функциям](custom-functions-requirement-sets.md)
* [Рекомендации по именованию](custom-functions-naming.md)
* [Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями](make-custom-functions-compatible-with-xll-udf.md)
