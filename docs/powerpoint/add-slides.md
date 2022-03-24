---
title: Добавление и удаление слайдов в PowerPoint
description: Узнайте, как добавлять и удалять слайды и указать мастер и макет новых слайдов.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: b14323a13332f2b1c9e26991c2446549ff78e745
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747008"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Добавление и удаление слайдов в PowerPoint

Надстройка PowerPoint добавить слайды в презентацию и дополнительно указать, какой мастер слайда и макет мастера используется для нового слайда. Надстройка также может удалять слайды.

API для добавления слайдов в основном используются в сценариях, в которых коды мастеров слайдов и макеты в презентации известны во время кодирования или могут быть найдены в источнике данных во время запуска. В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, имена или изображения мастеров слайдов и макетов) с ID-кодами мастеров слайдов и макетов. API также можно использовать в сценариях, где пользователь может вставлять слайды с использованием мастера слайдов по умолчанию и макета по умолчанию, а также в сценариях, в которых пользователь может выбрать существующий слайд и создать новый с тем же мастером слайда и макетом (но не с одним и тем же контентом). [Дополнительные сведения об этом](#select-which-slide-master-and-layout-to-use) см. в подборке мастера слайдов и макета.

## <a name="add-a-slide-with-slidecollectionadd"></a>Добавление слайда с помощью SlideCollection.add

Добавьте слайды [методом SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) . Ниже приводится простой пример, в котором добавляется слайд, использующий мастер слайдов презентации по умолчанию и первый макет этого мастера. Метод всегда добавляет новые слайды в конце презентации. Ниже приведен пример.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Выберите, какой мастер слайда и макет использовать

Используйте параметр [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) , чтобы контролировать, какой мастер слайда используется для нового слайда и какой макет используется в мастере. Ниже приведен пример. Вот что нужно знать об этом коде:

- Вы можете включить либо оба свойства `AddSlideOptions` объекта.
- Если используются оба свойства, указанный макет должен принадлежать указанному мастеру или ошибка будет выброшена.
- Если свойство `masterId` не присутствует (или его значение — пустая строка), `layoutId` используется мастер слайда по умолчанию и должен быть макет этого мастера слайдов.
- Мастер слайдов по умолчанию — это мастер слайдов, используемый последним слайдом в презентации. (В необычном случае, когда в настоящее время в презентации нет слайдов, мастер слайдов по умолчанию является первым мастером слайдов в презентации.)
- Если свойство `layoutId` не присутствует (или его значение — пустая строка), `masterId` используется первый макет мастера, заданный объектом.
- Оба свойства являются строками одной из трех возможных форм: ***nnnnnnnnnnnn*#**, **#* mmmmmmmmm***, или **_nnnnnnmmm_#****, где *nnnnnnnn* — это ID мастера или макета (обычно 10 цифр) и *ммм* — это код создания мастера или макета (обычно 6 — 10 цифр). Некоторые примеры , `2147483690#2908289500`и `2147483690#``#2908289500`.

```javascript
async function addSlide() {
    await PowerPoint.run(async function(context) {
        context.presentation.slides.add({
            slideMasterId: "2147483690#2908289500",
            layoutId: "2147483691#2499880"
        });
    
        await context.sync();
    });
}
```

Нет практических способов, чтобы пользователи могли обнаружить ID или создание ID мастера слайда или макета. По этой причине параметр `AddSlideOptions` можно использовать только в том случае, если вы знаете коды во время кодирования или ваша надстройка может обнаружить их во время работы. Так как нельзя ожидать, что пользователи будут запоминать ID, вам также потребуется способ, позволяющий пользователю выбрать слайды, возможно по имени или по изображению, а затем соотнести каждое название или изображение с ИД слайда.

Соответственно, этот `AddSlideOptions` параметр используется в основном в сценариях, в которых надстройка предназначена для работы с определенным набором мастеров слайдов и макетов, имена которых известны. В таком сценарии либо вы, либо клиент должны создать и сохранить источник данных, который сопоставляет критерий выбора (например, мастер слайдов и имена макетов или изображения) с соответствующими ID или кодами создания.

#### <a name="have-the-user-choose-a-matching-slide"></a>Чтобы пользователь выбрал совпадающий слайд

Если надстройка может использоваться в сценариях, в которых новый слайд должен использовать одно и то же сочетание мастера слайда и макета, используемого существующим слайдом, то надстройка может (1) подсказыть пользователю выбрать слайд и (2) прочитать ID мастера слайда и макет. В следующих действиях покажите, как читать ID и добавлять слайд с мастером и макетом.

1. Создайте метод, чтобы получить индекс выбранного слайда. Ниже приведен пример. Вот что нужно знать об этом коде:

    - Он использует [метод Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) общего API JavaScript.
    - Вызов встроен `getSelectedDataAsync` в функцию возврата обещаний. Дополнительные сведения о том, почему и как это сделать, см. в ссылке [Wrap Common API в функциях возврата обещаний](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions).
    - `getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов. В этом сценарии пользователь выбрал только один, поэтому код получает первый (0-й) слайд, который является единственным выбранным.
    - Значение `index` слайда — это 1-основанное значение, что пользователь видит рядом со слайдом в области эскизов.

    ```javascript
    function getSelectedSlideIndex() {
        return new OfficeExtension.Promise<number>(function(resolve, reject) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function(asyncResult) {
                try {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        reject(console.error(asyncResult.error.message));
                    } else {
                        resolve(asyncResult.value.slides[0].index);
                    }
                } 
                catch (error) {
                    reject(console.log(error));
                }
            });
        });
    }
    ```

2. Вызов новой функции [в PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) основной функции, которая добавляет слайд. Ниже приведен пример.

    ```javascript
    async function addSlideWithMatchingLayout() {
        await PowerPoint.run(async function(context) {
    
            let selectedSlideIndex = await getSelectedSlideIndex();
        
            // Decrement the index because the value returned by getSelectedSlideIndex()
            // is 1-based, but SlideCollection.getItemAt() is 0-based.
            const realSlideIndex = selectedSlideIndex - 1;
            const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");
        
            await context.sync();
        
            context.presentation.slides.add({
                slideMasterId: selectedSlide.slideMaster.id,
                layoutId: selectedSlide.layout.id
            });
        
            await context.sync();
        });
    }
    ```

## <a name="delete-slides"></a>Удаление слайдов

Удалите слайд, получив ссылку на объект [Slide](/javascript/api/powerpoint/powerpoint.slide) , который представляет слайд, и позвоните по методу `Slide.delete` . Ниже приводится пример удаления 4-го слайда.

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
