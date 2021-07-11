---
title: Вставка слайдов в PowerPoint презентации
description: Узнайте, как вставить слайды из одной презентации в другую.
ms.date: 03/07/2021
localization_priority: Normal
ms.openlocfilehash: 9b106e8940e7b0f19678e0467d8e900ffecd9438
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348785"
---
# <a name="insert-slides-in-a-powerpoint-presentation"></a>Вставка слайдов в PowerPoint презентации

Надстройка PowerPoint может вставлять слайды из одной презентации в текущую презентацию с помощью PowerPoint библиотеки JavaScript, определенной для приложений. Вы можете контролировать, будут ли вставлены слайды сохранять форматирование исходных презентаций или форматирование целевой презентации.

API вставки слайдов в основном используются в сценариях шаблонов презентаций: существует небольшое количество известных презентаций, которые служат пулами слайдов, которые могут быть вставлены надстройкой. В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, заголовки слайдов или изображения) с кодами слайдов. API также можно использовать в сценариях, в которых пользователь может вставлять слайды из любой произвольной  презентации, но в этом случае пользователь фактически ограничивается вставкой всех слайдов из исходных презентаций. Дополнительные [сведения об этом](#selecting-which-slides-to-insert) см. в подборе слайдов, которые необходимо вставить.

Существует два шага к вставке слайдов из одной презентации в другую.

1. Преобразование файла исходных презентаций (.pptx) в строку с форматом base64.
1. Используйте `insertSlidesFromBase64` метод, чтобы вставить один или несколько слайдов из файла base64 в текущую презентацию.

## <a name="convert-the-source-presentation-to-base64"></a>Преобразование исходных презентаций в base64

Существует множество способов преобразования файла в base64. Язык программирования и библиотека, которые вы используете, и преобразование на серверной стороне надстройки или клиентской стороне определяется вашим сценарием. Чаще всего преобразование в JavaScript будет происходить с клиентской стороны с помощью объекта [FileReader.](https://developer.mozilla.org/docs/Web/API/FileReader) В следующем примере показана эта практика.

1. Начните с получения ссылки на исходный PowerPoint файл. В этом примере мы будем использовать управление типом, чтобы побудить пользователя `<input>` `file` выбрать файл. Добавьте следующую разметку на страницу надстройки.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Эта разметка добавляет пользовательский интерфейс на следующий скриншот страницы.

    ![Снимок экрана, показывающий элемент управления вводом типа HTML-файлов, предшествующего инструкции по чтению предложения "Выберите презентацию PowerPoint, из которой вставить слайды". Управление состоит из кнопки с меткой "Выберите файл", за которой следует предложение "Нет выбранного файла".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Существует множество других способов получения PowerPoint файла. Например, если файл хранится на OneDrive или SharePoint, для его Graph Microsoft. Дополнительные сведения см. в материалах [Working with files in Microsoft Graph](/graph/api/resources/onedrive) и Access Files with Microsoft [Graph.](/learn/modules/msgraph-access-file-data/)

2. Добавьте следующий код в JavaScript надстройки, чтобы назначить функцию событию управления `change` входом. (Вы создаете `storeFileAsBase64` функцию на следующем шаге.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Добавьте в него указанный ниже код. Обратите внимание на следующие аспекты этого кода.

    - Метод `reader.readAsDataURL` преобразует файл в base64 и сохраняет его в `reader.result` свойстве. Когда метод завершается, он запускает `onload` обработник событий.
    - Обработник событий отделяет метаданные от закодированного файла и сохраняет закодированную строку `onload` в глобальной переменной.
    - Строка с кодом base64 хранится глобально, так как она будет считываться другой функцией, которую вы создаете на более позднем этапе.

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

## <a name="insert-slides-with-insertslidesfrombase64"></a>Вставка слайдов со вставкамиSlidesFromBase64

Ваша надстройка вставляет слайды из другого PowerPoint в текущую презентацию с помощью метода [Presentation.insertSlidesFromBase64.](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) Ниже приводится простой пример, в котором все слайды из презентации источника вставляются в начале текущей презентации, а вставленные слайды держат форматирование исходных файлов. Обратите внимание, что это глобальная переменная, которая содержит базовую версию файла PowerPoint `chosenFileBase64` презентации.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Вы можете управлять некоторыми аспектами результата вставки, в том числе с помощью вставки слайдов и получения источника или целевого форматирования, передав объект [InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра `insertSlidesFromBase64` . Ниже приведен пример. Вот что нужно знать об этом коде:

- Существует два возможных значения `formatting` свойства: "UseDestinationTheme" и "KeepSourceFormatting". Необязательный, вы можете использовать `InsertSlideFormatting` enum, (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).
- Функция будет вставлять слайды из презентации источника сразу после слайда, указанного `targetSlideId` свойством. Значение этого свойства — строка из трех возможных форм: ***nnn*#**, * *#* mmmmmmmmm***, или **_nnn_ #* mmmmmmmmm***, где nnn — это *ID* слайда (обычно 3 цифры), а *mmmmmmmmmmm* — это код создания слайда (обычно 9 цифр). Некоторые примеры , `267#763315295` `267#` и `#763315295` .

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

Конечно, во время кодирования обычно не будет знать ID или код создания целевого слайда. Чаще всего надстройка будет просить пользователей выбрать целевой слайд. В следующих действиях покажите, как получить ***nnn*#** ID выбранного в настоящее время слайда и использовать его в качестве целевого слайда.

1. Создайте функцию, которая получает ID выбранного в настоящее время слайда с помощью метода [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript. Ниже приведен пример. Обратите внимание, что вызов `getSelectedDataAsync` встроен в функцию возврата обещаний. Дополнительные сведения о том, почему и как это сделать, см. в Common-APIs wrap Common-APIs в функциях возврата [обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)

 
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

1. Вызов новой функции [PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции и передать возвращаемую (сопутованую символу #) ID в качестве значения свойства `targetSlideId` `InsertSlideOptions` параметра. Ниже приведен пример.

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

Вы также можете использовать параметр [InsertSlideOptions,](/javascript/api/powerpoint/powerpoint.insertslideoptions) чтобы контролировать, какие слайды из презентации источника вставляются. Это необходимо, назначив свойству массив слайд-кодов исходных `sourceSlideIds` презентаций. Ниже приводится пример, в который вставляется четыре слайда. Обратите внимание, что каждая строка в массиве должна следовать тем или иным шаблонам, используемым для `targetSlideId` свойства.

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

Нет практического способа, чтобы пользователи могли обнаружить ID или код создания слайда в презентации источника. По этой причине свойство можно использовать только в том случае, если вы знаете исходные коды во время кодирования или ваша надстройка может получить их в период работы из источника `sourceSlideIds` данных. Поскольку нельзя ожидать, что пользователи будут запоминать слайд-ИД, вам также необходим способ, позволяющий пользователю выбрать слайды, возможно, по названию или по изображению, а затем соотнести каждое название или изображение с ИД слайда.

Соответственно, свойство используется в основном в сценариях шаблонов презентаций: надстройка предназначена для работы с определенным набором презентаций, которые служат пулами слайдов, которые можно `sourceSlideIds` вставить. В таком сценарии либо вы, либо клиент должны создавать и поддерживать источник данных, который сопоставляет критерий выбора (например, заголовки или изображения) с кодами слайдов или кодами создания слайдов, которые были построены из набора возможных исходных презентаций.
