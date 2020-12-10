---
title: Вставка и удаление слайдов в презентации PowerPoint
description: Сведения о том, как вставлять слайды из одной презентации в другую и удалять слайды.
ms.date: 12/04/2020
localization_priority: Normal
ms.openlocfilehash: ceb78054a95ac4b26bd71f79a086a00e3dce5278
ms.sourcegitcommit: cba180ae712d88d8d9ec417b4d1c7112cd8fdd17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/09/2020
ms.locfileid: "49613711"
---
# <a name="insert-and-delete-slides-in-a-powerpoint-presentation-preview"></a>Вставка и удаление слайдов в презентации PowerPoint (Предварительная версия)

Надстройка PowerPoint позволяет вставлять слайды из одной презентации в текущую, используя библиотеку JavaScript, зависящую от приложения PowerPoint. Вы можете указать, следует ли вставлять вставленные слайды в исходную презентацию или форматировать целевую презентацию. Вы также можете удалять слайды из презентации.

[!include[General preview API prerequisites](../includes/using-preview-apis-host.md)]

API вставки слайдов в основном используются в сценариях презентации: существует небольшое количество известных презентаций, которые могут быть вставлены надстройкой в виде пулов слайдов. В этом сценарии либо вы, либо клиент должны создать и поддерживать источник данных, который соответствует условию выбора (например, заголовки слайдов или изображения) с идентификаторами слайдов. API также можно использовать в сценариях, где пользователь может вставлять слайды из произвольной презентации, но в этом сценарии пользователь практически ограничен вставкой *всех* слайдов из исходной презентации. Дополнительные сведения об этом [можно узнать в разделе Выбор слайдов для вставки](#selecting-which-slides-to-insert) .

Вставить слайды из одной презентации в другую можно двумя шагами.

1. Преобразование исходного файла презентации (PPTX) в строку в формате Base64.
1. Используйте `insertSlidesFromBase64` метод, чтобы вставить один или несколько слайдов из файла Base64 в текущую презентацию.

## <a name="convert-the-source-presentation-to-base64"></a>Преобразование исходной презентации в формат Base64

Существует множество способов преобразования файла в Base64. Выбор используемого языка программирования и библиотеки, а также необходимость преобразования на стороне сервера надстройки или на стороне клиента определяется сценарием. Чаще всего преобразование в JavaScript на стороне клиента выполняется с помощью объекта [FileReader браузером](https://developer.mozilla.org/docs/Web/API/FileReader) . В приведенном ниже примере показана эта практика.

1. Сначала получите ссылку на исходный файл PowerPoint. В этом примере мы будем использовать `<input>` элемент управления типа, `file` чтобы предлагать пользователю выбрать файл. Добавьте указанную ниже разметку на страницу надстройки.

    ```html
    <section>
        <p>Select a PowerPoint presentation from which to insert slides</p>
        <form>
            <input type="file" id="file" />
        </form>
    </section>
    ```

    Эта разметка добавляет пользовательский интерфейс на страницу на следующем снимке экрана:

    ![Снимок экрана с элементом управления вводом типа HTML-файла, которому предшествует пояснительное предложение чтение "Выбор презентации PowerPoint, из которой нужно вставить слайды". Элемент управления состоит из кнопки с надписью "выберите файл", за которой следует предложение "файл не выбран".](../images/powerpoint-html-file-input-control.png)

    > [!NOTE]
    > Существует множество других способов получения файла PowerPoint. Например, если файл хранится в OneDrive или SharePoint, вы можете скачать его с помощью Microsoft Graph. Дополнительные сведения см. [в статье работа с файлами в Microsoft Graph](/graph/api/resources/onedrive) и [доступ к файлам с помощью Microsoft Graph](/learn/modules/msgraph-access-file-data/).

2. Добавьте следующий код в JavaScript надстройки, чтобы назначить функцию для события элемента управления вводом `change` . (Вы создадите `storeFileAsBase64` функцию на следующем шаге.)

    ```javascript
    $("#file").change(storeFileAsBase64);
    ```

3. Добавьте в него указанный ниже код. Обратите внимание на следующие особенности этого кода:

    - `reader.readAsDataURL`Метод преобразует файл в формат Base64 и сохраняет его в `reader.result` свойстве. После выполнения метода он запускает `onload` обработчик событий.
    - `onload`Обработчик событий удаляет метаданные из зашифрованного файла и сохраняет закодированную строку в глобальной переменной.
    - Строка в кодировке Base64 хранится глобально, так как она будет прочитана другой функцией, созданной на более позднем этапе.

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

Надстройка вставляет слайды из другой презентации PowerPoint в текущую презентацию с помощью метода [Presentation. insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#insertslidesfrombase64-base64file--options-) . Ниже приведен простой пример, в котором все слайды исходной презентации вставляются в начало текущей презентации, а вставленные слайды сохраняются в формате исходного файла. Обратите внимание, что `chosenFileBase64` это глобальная переменная, содержащая версию файла презентации PowerPoint в кодировке Base64.

```javascript
async function insertAllSlides() {
  await PowerPoint.run(async function(context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}
```

Вы можете управлять некоторыми аспектами вставки, включая место вставки слайдов и способ получения исходного или целевого форматирования, путем передачи объекта [инсертслидеоптионс](/javascript/api/powerpoint/powerpoint.insertslideoptions) в качестве второго параметра `insertSlidesFromBase64` . Ниже приведен пример. Вот что нужно знать об этом коде:

- Свойство имеет два возможных значения `formatting` : "уседестинатионсеме" и "кипсаурцеформаттинг". При необходимости можно использовать `InsertSlideFormatting` Перечисление (например, `PowerPoint.InsertSlideFormatting.useDestinationTheme` ).
- Функция вставит слайды из исходной презентации сразу же после слайда, указанного `targetSlideId` свойством. Значение этого свойства является строкой одной из трех возможных форм: ***nnn * #**, * *#* ммммммммм * * * или **_nnn_ #* ммммммммм * * *, где *nnn* — идентификатор слайда (обычно 3 цифры), а *ммммммммм* — идентификатор создания слайда (обычно 9 цифры). Некоторые примеры: `267#763315295` , `267#` и `#763315295` .

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

Конечно, вы, как правило, не узнаете на момент кодирования идентификатор или идентификатор создания целевого слайда. Чаще всего надстройка запрашивает у пользователей Выбор целевого слайда. В следующей процедуре показано, как получить идентификатор ***nnn * #** выбранного слайда и использовать его в качестве целевого слайда.

1. Создайте функцию, которая получает идентификатор текущего выбранного слайда с помощью метода [Office.context.docумент. GetSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript. Ниже приведен пример. Обратите внимание, что вызов `getSelectedDataAsync` внедряется в функцию, возвращающую обещание. Дополнительные сведения о том, почему и как это сделать, можно узнать [в статье Wrap Common-APIs в функциях, возвращающих обещаний](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).

 
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

1. Вызовите новую функцию в [PowerPoint. Run ()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции и передайте возвращенный идентификатор (сцепленный с символом "#") в качестве значения `targetSlideId` свойства этого `InsertSlideOptions` параметра. Ниже приведен пример.

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

Кроме того, можно использовать параметр [инсертслидеоптионс](/javascript/api/powerpoint/powerpoint.insertslideoptions) для управления вставкой слайдов из исходной презентации. Для этого необходимо назначить свойству массив идентификаторов слайдов исходной презентации `sourceSlideIds` . Ниже приведен пример вставки четырех слайдов. Обратите внимание, что каждая строка в массиве должна соответствовать одному или другому шаблону, используемому для `targetSlideId` Свойства.

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
> Слайды будут вставлены в один и тот же относительный порядок, в котором они отображаются в исходной презентации независимо от того, в каком порядке они отображаются в массиве.

Не существует практического способа, с помощью которого пользователи могут обнаружить идентификатор или идентификатор создания слайда в исходной презентации. По этой причине это свойство можно использовать только в том `sourceSlideIds` случае, если вы знаете идентификаторы источника на момент написания кода или надстройка может получить их в среде выполнения из некоторого источника данных. Так как пользователи не могут запоминать идентификаторы слайдов, также необходим способ, позволяющий пользователю выбирать слайды, возможно, по названию или изображению, а затем сопоставлять каждое название или изображение с ИДЕНТИФИКАТОРом слайда.

Соответственно, `sourceSlideIds` свойство используется в сценариях презентации: надстройка разработана для работы с определенным набором презентаций, которые используются в качестве пулов слайдов, которые можно вставить. В этом сценарии либо вы, либо клиент должны создать и поддерживать источник данных, который соответствует условию выбора (например, заголовки или изображения) с идентификаторами слайдов или идентификаторами создания слайдов, созданными из набора возможных исходных презентаций.

## <a name="delete-slides"></a>Удаление слайдов

Вы можете удалить слайд, получив ссылку на объект [слайда](/javascript/api/powerpoint/powerpoint.slide) , представляющий слайд, и вызовите `Slide.delete` метод. Ниже приведен пример, в котором удаляется четвертый слайд.

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
