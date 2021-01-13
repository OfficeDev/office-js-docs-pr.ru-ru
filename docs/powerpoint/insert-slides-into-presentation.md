---
title: Вставка и удаление слайдов в презентации PowerPoint
description: Узнайте, как вставлять слайды из одной презентации в другую и как удалять слайды.
ms.date: 01/08/2021
localization_priority: Normal
ms.openlocfilehash: a9a4b2efd1e970d9c45885f9a17046bec4de7e72
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839721"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation"></a>Вставка и удаление слайдов в презентации PowerPoint

Надстройка PowerPoint может вставлять слайды из одной презентации в текущую презентацию с помощью библиотеки JavaScript для конкретного приложения PowerPoint. Можно контролировать, будут ли вставлены слайды сохранять форматирование исходных презентаций или форматирование целевой презентации. Вы также можете удалить слайды из презентации.

API вставки слайдов в основном используются в сценариях шаблонов презентаций: существует небольшое количество известных презентаций, которые служат в качестве пулов слайдов, которые могут быть вставлены надстройкой. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки слайдов или изображения) с кодами слайдов. API также можно использовать в сценариях, где пользователь может вставлять слайды из любой произвольной презентации, но в этом сценарии пользователь фактически ограничивается вставкой всех слайдов из презентации источника.  Дополнительные сведения об этом [см.](#selecting-which-slides-to-insert) в под вопросе "Выбор слайдов для вставки".

Вставка слайдов из одной презентации в другую состоит из двух этапов.

1. Преобразуем исходный файл презентации (PPTX) в строку в формате base64.
1. Используйте этот метод, чтобы вставить один или несколько слайдов из `insertSlidesFromBase64` файла base64 в текущую презентацию.

## <a name="convert-the-source-presentation-to-base64"></a>Преобразование исходных презентаций в base64

Существует множество способов преобразования файла в base64. Язык программирования и библиотека, которые вы используете, и будет ли преобразование на стороне сервера надстройки или на стороне клиента определяется вашим сценарием. Чаще всего преобразование в JavaScript происходит на стороне клиента с помощью объекта [FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader) В следующем примере показано, как это сделать.

1. Начните с получения ссылки на исходный файл PowerPoint. В этом примере мы будем использовать тип для запроса на выбор `<input>` `file` файла пользователем. Добавьте следующую разметку на страницу надстройки.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Эта разметка добавляет пользовательский интерфейс на следующий снимок экрана на страницу:

    ![Screenshot showing an HTML file type input control preceded by an instructional sentence reading "Select a PowerPoint presentation from which to insert slides". Этот контроль состоит из кнопки "Выбрать файл", за которой следует предложение "Файл не выбран".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Существует множество других способов получения файла PowerPoint. Например, если файл хранится в OneDrive или SharePoint, его можно скачать с помощью Microsoft Graph. Дополнительные сведения см. [в работе с файлами в Microsoft Graph](/graph/api/resources/onedrive) и access Files с помощью Microsoft [Graph.](/learn/modules/msgraph-access-file-data/)

2. Добавьте следующий код в Код JavaScript надстройки, чтобы назначить функцию событию входного `change` управления. (Функция `storeFileAsBase64` создается на следующем этапе.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Добавьте в него указанный ниже код. Обратите внимание на следующие вопросы об этом коде:

    - Метод `reader.readAsDataURL` преобразует файл в base64 и сохраняет его в `reader.result` свойстве. После завершения работы метода запускается `onload` обработок события.
    - Обработник событий отключит метаданные закодированного файла и сохраняет закодированную строку `onload` в глобальной переменной.
    - Строка в кодировке base64 хранится глобально, так как она будет считываться другой функцией, которая будет создаваться на одном из последующих этапов.

    ```javascript
    let chosenFileBase64;

    async function storeFileAsBase64() {
        const reader = new FileReader();

        reader.onload = async (event) => {
            const startIndex = reader.result.toString().indexOf("base64,");
            const copyBase64 = reader.result.toString().substr(startIndex + 7);

            chosenFileBase64 = copyBase64;
        };

        const myFile = document.getElementById("file") as HTMLInputElement;
        reader.readAsDataURL(myFile.files[0]);
    }
    ```

## <a name="insert-slides-with-insertslidesfrombase64"></a>Вставка слайдов с помощью insertSlidesFromBase64

Надстройка вставляет слайды из другой презентации PowerPoint в текущую презентацию с помощью метода [Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) Ниже приводится простой пример, в котором все слайды из презентации источника вставляются в начале текущей презентации, а вставляемые слайды сохранят форматирование исходных файлов. Обратите внимание, что это глобальная переменная, которая содержит версию файла презентации PowerPoint в кодировке `chosenFileBase64` base64.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Вы можете управлять некоторыми аспектами результата вставки, включая место вставки слайдов и то, получают ли они форматирование источника или цели, передав объект [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра в `insertSlidesFromBase64` . Ниже приведен пример. Вот что нужно знать об этом коде:

- Свойство может иметь два возможных `formatting` значения: UseDestinationTheme и KeepSourceFormatting. При желании можно использовать `InsertSlideFormatting` это enum (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` )..
- Функция вставляет слайды из презентации источника сразу после слайда, указанного `targetSlideId` свойством. Значение этого свойства является строкой одной из трех возможных форм: ***nnn*#**, * *#* mmm***, или **_nnn_ #* mmm***, где *nnn* — это  ИД слайда (обычно 3 цифры), а ммм — это код создания слайда (обычно 9 цифр). Некоторые примеры: `267#763315295` `267#` , и `#763315295` .

```javascript
async function insertSlidesDestinationFormatting() {
  await PowerPoint.run(async function(context) {
    context.presentation
    .insertSlidesFromBase64(chosenFileBase64,
                            {
                                formatting: "UseDestinationTheme",
                                targetSlideId: "267#"
                            }
                          );
    await context.sync();
  });
}
```

Конечно, во время написания кода вы, как правило, не знаете ИД или ид создания целевого слайда. Чаще всего надстройка просит пользователей выбрать целевой слайд. Далее покажите, как получить ***nnn*#** ИД выбранного слайда и использовать его в качестве целевого слайда.

1. Создайте функцию, которая получает ИД выбранного слайда с помощью метода [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript. Ниже приведен пример. Обратите внимание, что вызов `getSelectedDataAsync` внедрен в функцию возврата обещания. Дополнительные сведения о том, почему и как это сделать, см. в Common-APIs wrap в функциях возврата [обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)

 
    ```javascript
    function getSelectedSlideID() {
      return new OfficeExtension.Promise<string>(function (resolve, reject) {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
          try {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              reject(console.error(asyncResult.error.message));
            } else {
              resolve(asyncResult.value.slides[0].id);
            }
          }
          catch (error) {
            reject(console.log(error));
          }
        });
      })
    }
    ```

1. Вызовите новую функцию внутри [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) основной функции и передайте возвращаемую ид (вмещенную символом "#") в качестве значения свойства `targetSlideId` `InsertSlideOptions` параметра. Ниже приведен пример.

    ```javascript
    async function insertAfterSelectedSlide() {
        await PowerPoint.run(async function(context) {

            const selectedSlideID = await getSelectedSlideID();

            context.presentation.insertSlidesFromBase64(chosenFileBase64, {
                formatting: "UseDestinationTheme",
                targetSlideId: selectedSlideID + "#"
            });

            await context.sync();
        });
    }
    ```

### <a name="selecting-which-slides-to-insert"></a>Выбор слайдов для вставки

Вы также можете использовать параметр [InsertSlideOptions,](/javascript/api/powerpoint/powerpoint.insertslideoptions) чтобы контролировать, какие слайды из презентации источника вставляются. Для этого свойству назначается массив кодов слайдов презентации `sourceSlideIds` источника. Ниже приводится пример вставки четырех слайдов. Обратите внимание, что каждая строка в массиве должна следовать одному или одному из шаблонов, используемых для `targetSlideId` свойства.

```javascript
async function insertAfterSelectedSlide() {
    await PowerPoint.run(async function(context) {
        const selectedSlideID = await getSelectedSlideID();
        context.presentation.insertSlidesFromBase64(chosenFileBase64, {
            formatting: "UseDestinationTheme",
            targetSlideId: selectedSlideID + "#",
            sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"]
        });

        await context.sync();
    });
}
```

> [!NOTE]
> Слайды будут вставлены в том же относительном порядке, в котором они отображаются в презентации источника, независимо от порядка, в котором они отображаются в массиве.

Не существует практического способа обнаружения пользователем ИД или ид создания слайда в презентации источника. По этой причине свойство можно использовать, только если вы знаете исходные коды во время кодирования или надстройка может получить их во время работы из некоторого `sourceSlideIds` источника данных. Так как пользователи не могут запоминать ид слайдов, вам также потребуется способ, позволяющий пользователю выбирать слайды, например по названию или изображению, а затем соотносить каждый заголовок или изображение с ид слайда.

Соответственно, свойство в основном используется в сценариях шаблонов презентаций: надстройка предназначена для работы с определенным набором презентаций, которые выступать в качестве пулов слайдов, которые можно `sourceSlideIds` вставить. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки или изображения) с кодами слайдов или идами создания слайдов, которые были сконструированы из набора возможных исходных презентаций.

## <a name="delete-slides"></a>Удаление слайдов

Вы можете удалить слайд, получив ссылку на объект [Slide,](/javascript/api/powerpoint/powerpoint.slide) который представляет слайд, и вызовите `Slide.delete` метод. Ниже приводится пример удаления 4-го слайда.

```javascript
async function deleteSlide() {
  await PowerPoint.run(async function(context) {

    // The slide index is zero-based. 
    const slide = context.presentation.slides.getItemAt(3);
    slide.delete();
    await context.sync();
  });
}
```
