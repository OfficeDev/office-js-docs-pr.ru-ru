---
title: Вставка слайдов в презентацию PowerPoint
description: Узнайте, как вставлять слайды из одной презентации в другую.
ms.date: 03/07/2021
ms.localizationpriority: medium
ms.openlocfilehash: a31933de4272634394dc6c36aafa973c41265471
ms.sourcegitcommit: 54a7dc07e5f31dd5111e4efee3e85b4643c4bef5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2022
ms.locfileid: "67857573"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Вставка слайдов в презентацию PowerPoint

Надстройка PowerPoint может вставлять слайды из одной презентации в текущую презентацию с помощью библиотеки JavaScript для приложения PowerPoint. Вы можете контролировать, будут ли вставленные слайды сохранять форматирование исходной презентации или форматирование целевой презентации.

API-интерфейсы вставки слайдов в основном используются в сценариях шаблонов презентаций. Существует небольшое количество известных презентаций, которые служат пулами слайдов, которые могут быть вставлены надстройкой. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки слайдов или изображения) с идентификаторами слайдов. API-интерфейсы также можно использовать в сценариях, в которых пользователь может вставлять слайды из любой произвольной презентации, но в этом сценарии пользователь фактически ограничивается  вставкой всех слайдов из исходной презентации. [Дополнительные сведения об этом см](#selecting-which-slides-to-insert). в разделе "Выбор слайдов для вставки".

Вставить слайды из одной презентации в другую можно двумя способами.

1. Преобразуйте исходный файл презентации (.pptx) в строку в формате base64.
1. Используйте этот `insertSlidesFromBase64` метод, чтобы вставить один или несколько слайдов из файла Base64 в текущую презентацию.

## <a name="convert-the-source-presentation-to-base64"></a>Преобразование исходной презентации в base64

Существует множество способов преобразования файла в base64. Используемый язык программирования и библиотека, а также необходимость преобразования на стороне сервера надстройки или на стороне клиента определяются вашим сценарием. Чаще всего преобразование выполняется в JavaScript на стороне клиента с помощью объекта [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) . В следующем примере показано, как это сделать.

1. Начните с получения ссылки на исходный файл PowerPoint. В этом примере мы будем использовать элемент `<input>` управления типа, `file` чтобы предложить пользователю выбрать файл. Добавьте следующую разметку на страницу надстройки.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Эта разметка добавляет пользовательский интерфейс на следующем снимке экрана на страницу.

    ![Снимок экрана: элемент управления вводом типа HTML-файла, предшествующий инструкционном предложению с текстом "Выберите презентацию PowerPoint, из которой нужно вставить слайды". Элемент управления состоит из кнопки "Выбрать файл", за которой следует предложение "Файл не выбран".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Существует множество других способов получить файл PowerPoint. Например, если файл хранится в OneDrive или SharePoint, его можно скачать с помощью Microsoft Graph. Дополнительные сведения см. в [статье "Работа с файлами в Microsoft Graph](/graph/api/resources/onedrive) и [access Files с помощью Microsoft Graph"](/training/modules/msgraph-access-file-data/).

2. Добавьте следующий код в Код JavaScript надстройки, чтобы назначить функцию событию элемента управления вводом `change` . (Вы создайте функцию `storeFileAsBase64` на следующем шаге.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Добавьте в него указанный ниже код. Обратите внимание на указанные ниже аспекты этого кода.

    - Метод `reader.readAsDataURL` преобразует файл в base64 и сохраняет его в свойстве `reader.result` . После завершения работы метода он активирует обработчик `onload` событий.
    - Обработчик `onload` событий удаляет метаданные закодированного файла и сохраняет закодированную строку в глобальной переменной.
    - Строка в кодировке Base64 хранится глобально, так как она будет считываться другой функцией, которая будет создаваться на следующем шаге.

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

Ваша надстройка вставляет слайды из другой презентации PowerPoint в текущую презентацию с помощью метода [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1)) . Ниже приведен простой пример, в котором все слайды из исходной презентации вставляются в начале текущей презентации, а вставленные слайды хранят форматирование исходного файла. Обратите внимание, `chosenFileBase64` что это глобальная переменная, содержащая версию файла презентации PowerPoint в кодировке Base64.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Вы можете управлять некоторыми аспектами результата вставки, включая место вставки слайдов и получение исходного или целевого форматирования, передав объект [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) `insertSlidesFromBase64`в качестве второго параметра. Ниже приведен пример. Вот что нужно знать об этом коде:

- Существует два возможных значения `formatting` свойства: UseDestinationTheme и KeepSourceFormatting. При необходимости можно использовать перечисление `InsertSlideFormatting` (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme`).
- Функция вставляет слайды из исходной презентации сразу после слайда, указанного свойством `targetSlideId` . Значение этого свойства представляет собой строку из трех возможных форм: ***nnn*#**, **#* mmmmmmmmm*** или **_nnn_#* mmmmmmmmm***, где *nnn* — это идентификатор слайда (обычно 3 цифры), а *mmmmmmmmm* — идентификатор создания слайда (обычно 9 цифр). Вот несколько примеров: `267#763315295`и `267#``#763315295`.

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

Конечно, во время написания кода идентификатор или идентификатор создания целевого слайда обычно не будут известны. Чаще всего надстройка запрашивает у пользователей выбор целевого слайда. Ниже показано, как получить идентификатор ***nnn*#** выбранного слайда и использовать его в качестве целевого слайда.

1. Создайте функцию, которая получает идентификатор выбранного слайда с помощью метода [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) общих API JavaScript. Ниже приведен пример. Обратите внимание, что вызов внедрен `getSelectedDataAsync` в функцию, возвращаемую обещанием. Дополнительные сведения о том, почему и как это сделать, см. Common-APIs в функциях, возвращаемых [обещанием](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

1. Вызовите новую функцию внутри [файла PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции и передайте возвращаемый идентификатор (объединенный с символом "#") `targetSlideId` `InsertSlideOptions` в качестве значения свойства параметра. Ниже приведен пример.

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

Вы также можете использовать параметр [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) для управления вставкой слайдов из исходной презентации. Для этого свойству назначается массив идентификаторов слайдов исходной `sourceSlideIds` презентации. Ниже приведен пример вставки четырех слайдов. Обратите внимание, что каждая строка в массиве должна соответствовать одному или нескольким шаблонам, используемым для `targetSlideId` свойства.

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
> Слайды будут вставлены в том же относительном порядке, в котором они отображаются в исходной презентации, независимо от порядка их отображения в массиве.

Нет практического способа, с помощью которого пользователи могут обнаружить идентификатор или идентификатор создания слайда в исходной презентации. По этой причине свойство можно использовать только в том случае, `sourceSlideIds` если во время кодирования вы знаете исходные идентификаторы или надстройка может получить их во время выполнения из некоторого источника данных. Так как пользователи не могут запоминать идентификаторы слайдов, вам также нужен способ, позволяющий пользователю выбирать слайды, например по заголовку или изображению, а затем сопоставлять каждый заголовок или изображение с идентификатором слайда.

Соответственно, это `sourceSlideIds` свойство в основном используется в сценариях шаблонов презентаций. Надстройка предназначена для работы с определенным набором презентаций, которые служат пулами слайдов, которые можно вставить. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки или изображения) с идентификаторами слайдов или идентификаторами создания слайдов, созданными на основе набора возможных исходных презентаций.
