---
title: Добавление и удаление слайдов в PowerPoint
description: Узнайте, как добавлять и удалять слайды, а также указывать образец и макет новых слайдов.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2cf22c18cf4089bab9091be3f4274f67974662a3
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958315"
---
# <a name="add-and-delete-slides-in-powerpoint"></a>Добавление и удаление слайдов в PowerPoint

Надстройка PowerPoint может добавлять слайды в презентацию и при необходимости указывать, какой образец слайдов и какой макет образца используется для нового слайда. Надстройка также может удалять слайды.

API для добавления слайдов в основном используются в сценариях, где идентификаторы образцов слайдов и макетов в презентации известны во время написания кода или могут быть найдены в источнике данных во время выполнения. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, имена или изображения образцов слайдов и макетов) с идентификаторами образцов слайдов и макетов. API-интерфейсы также можно использовать в сценариях, где пользователь может вставлять слайды, использующие образец слайдов по умолчанию и макет образца по умолчанию, а также в сценариях, где пользователь может выбрать существующий слайд и создать новый слайд с тем же образцом слайдов и макетом (но не тем же содержимым). [Дополнительные сведения об этом см](#select-which-slide-master-and-layout-to-use). в разделе "Выбор образца слайдов и макета".

## <a name="add-a-slide-with-slidecollectionadd"></a>Добавление слайда с помощью SlideCollection.add

Добавление слайдов с помощью [метода SlideCollection.add](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1)) . Ниже приведен простой пример, в котором добавляется слайд, использующий образец слайдов презентации по умолчанию и первый макет этого образца. Этот метод всегда добавляет новые слайды в конец презентации. Ниже приведен пример.

```javascript
async function addSlide() {
  await PowerPoint.run(async function(context) {
    context.presentation.slides.add();

    await context.sync();
  });
}
```

### <a name="select-which-slide-master-and-layout-to-use"></a>Выбор образца слайдов и макета для использования

Используйте параметр [AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions) для управления тем, какой образец слайдов используется для нового слайда и какой макет используется в образце. Ниже приведен пример. Вот что нужно знать об этом коде:

- Можно включить либо одно, либо оба свойства `AddSlideOptions` объекта.
- Если используются оба свойства, указанный макет должен принадлежать указанному главному объекту или возникает ошибка.
- Если свойство `masterId` отсутствует (или его значение является пустой строкой), `layoutId` используется образец слайдов по умолчанию и должен быть макет этого образца слайдов.
- Образец слайдов по умолчанию — это образец слайдов, используемый последним слайдом в презентации. (В необычном случае, когда в настоящее время в презентации нет слайдов, образец слайдов по умолчанию является первым образцом слайдов в презентации.)
- Если свойство `layoutId` отсутствует (или его значение является пустой строкой), используется первый макет образца, указанный объектом `masterId` .
- Оба свойства представляют собой строки одной из трех возможных форм: ***nnnnnnnnnn*#**, **#* mmmmmmmmm*** или **_nnnnnnnnnn_#* mmmmmmm***, где *nnnnnnnnnn* — это идентификатор образца или макета (обычно 10 цифр), а *mmmmmmmmm* — идентификатор создания образца или макета (обычно 6–10 цифр). Вот несколько примеров: `2147483690#2908289500`и `2147483690#``#2908289500`.

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

Нет практического способа, с помощью которого пользователи могут обнаружить идентификатор или идентификатор создания образца слайдов или макета. По этой причине этот `AddSlideOptions` параметр можно использовать только в том случае, если идентификаторы известны во время написания кода или надстройка может обнаружить их во время выполнения. Так как пользователи не могут запоминать идентификаторы, вам также нужен способ, позволяющий пользователю выбирать слайды по имени или изображению, а затем сопоставлять каждый заголовок или изображение с идентификатором слайда.

Соответственно, этот `AddSlideOptions` параметр в основном используется в сценариях, в которых надстройка предназначена для работы с определенным набором образцов слайдов и макетов, идентификаторы которых известны. В таком сценарии вам или клиенту необходимо создать и поддерживать источник данных, который сопоставляет критерий выбора (например, образец слайдов, имена макетов или изображения) с соответствующими идентификаторами или идентификаторами создания.

#### <a name="have-the-user-choose-a-matching-slide"></a>Выбор соответствующего слайда пользователем

Если надстройку можно использовать в сценариях, где новый слайд должен использовать то же сочетание образца слайдов и макета, что и существующий  слайд, надстройка может (1) предложить пользователю выбрать слайд и (2) прочитать идентификаторы образца слайдов и макета. Ниже показано, как прочитать идентификаторы и добавить слайд с соответствующим образцом и макетом.

1. Создайте функцию для получения индекса выбранного слайда. Ниже приведен пример. Вот что нужно знать об этом коде:

    - Он использует метод [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) общих API JavaScript.
    - Вызов внедряется `getSelectedDataAsync` в функцию, возвращаемую обещанием. Дополнительные сведения о том, почему и как это сделать, см. в статье "Упаковка общих [API-интерфейсов в](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions) функции, возвращающие обещание".
    - `getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов. В этом сценарии пользователь выбирает только один, поэтому код получает первый (0) слайд, который является единственным выбранным.
    - Значение `index` слайда — это значение на основе 1, которое пользователь видит рядом со слайдом в области эскизов.

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

2. Вызовите новую функцию в [файле PowerPoint.run()](/javascript/api/powerpoint#PowerPoint_run_batch_) главной функции, которая добавляет слайд. Ниже приведен пример.

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

Удалите слайд, используя ссылку на объект [Slide](/javascript/api/powerpoint/powerpoint.slide) , представляющий слайд, и вызовите `Slide.delete` метод. Ниже приведен пример удаления 4-го слайда.

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
