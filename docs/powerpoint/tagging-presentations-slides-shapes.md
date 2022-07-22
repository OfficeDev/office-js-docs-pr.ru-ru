---
title: Использование настраиваемых тегов для презентаций, слайдов и фигур в PowerPoint
description: Узнайте, как использовать теги для пользовательских метаданных о презентациях, слайдах и фигурах.
ms.date: 12/14/2021
ms.localizationpriority: medium
ms.openlocfilehash: a30beea56286437b1c69461534ca13912107cecf
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958904"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Использование настраиваемых тегов для презентаций, слайдов и фигур в PowerPoint

Надстройка может присоединять пользовательские метаданные в виде пар "ключ-значение" (теги) к презентациям, определенным слайдам и определенным фигурам на слайде.

Существует два основных сценария использования тегов:

- При применении к слайду или фигуре тег позволяет классифицирует объект для пакетной обработки. Например, предположим, что презентация содержит слайды, которые должны быть включены в презентации в регион "Восточная часть", но не в регион "Западная часть". Аналогичным образом существуют альтернативные слайды, которые должны отображаться только в западной части. Надстройка может создать `REGION` `East` тег с ключом и значением и применить его к слайдам, которые должны использоваться только на востоке. Значение тега устанавливается для `West` слайдов, которые должны отображаться только в регионе "Западная часть". Перед презентацией на востоке кнопка в надстройке запускает код, который проходит по всем слайдам, проверяя значение тега `REGION` . Слайды, на которых удаляется `West` область. Затем пользователь закрывает надстройку и запускает слайд-шоу.
- При применении к презентации тег фактически является настраиваемым свойством в документе презентации (аналогично [CustomProperty](/javascript/api/word/word.customproperty) в Word).

## <a name="tag-slides-and-shapes"></a>Добавление тегов к слайдам и фигурам

Тег — это пара "ключ-значение", `string` где значение всегда имеет тип и представлено [объектом Tag](/javascript/api/powerpoint/powerpoint.tag) . Каждый тип родительского объекта, например [Presentation](/javascript/api/powerpoint/powerpoint.presentation), [Slide](/javascript/api/powerpoint/powerpoint.slide) или [Shape](/javascript/api/powerpoint/powerpoint.shape) , `tags` имеет свойство [типа TagsCollection](/javascript/api/powerpoint/powerpoint.tagcollection).

### <a name="add-update-and-delete-tags"></a>Добавление, обновление и удаление тегов

Чтобы добавить тег к объекту, вызовите [метод TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1)) свойства родительского `tags` объекта. Следующий код добавляет два тега на первый слайд презентации. Вот что нужно знать об этом коде:

- Первый параметр метода — `add` это ключ в паре "ключ-значение".
- Второй параметр — это значение.
- Ключ имеет прописные буквы. `add` Это не является строго обязательным для метода, однако ключ всегда хранится в PowerPoint в верхнем регистре, а некоторые методы, связанные с тегами, требуют, чтобы ключ был выражен в верхнем регистре *, поэтому* рекомендуется всегда использовать верхний регистр в коде для ключа тега.

```javascript
async function addMultipleSlideTags() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}
```

Этот `add` метод также используется для обновления тега. Следующий код изменяет значение тега `PLANET` .

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

Чтобы удалить тег, вызовите `delete` метод родительского `TagsCollection` объекта и передайте ключ тега в качестве параметра. Пример см. в разделе ["Задание пользовательских метаданных в презентации"](#set-custom-metadata-on-the-presentation).

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Выборочная обработка слайдов и фигур с помощью тегов

Рассмотрим следующий сценарий: Компания Contoso Consulting предоставляет презентацию, которая будет отображаться для всех новых клиентов. Но некоторые слайды должны отображаться только для клиентов, оплатив состояние "Премиум". Перед показом презентации пользователям, не являмся клиентами ценовой категории "Премиум", они делают ее копию и удаляют слайды, которые должны видеть только пользователи уровня "Премиум". Надстройка позволяет Компании Contoso пометить слайды, которые предназначены для клиентов уровня "Премиум", и при необходимости удалить их. В следующем списке описаны основные шаги кодирования для создания этой функции.

1. Создайте функцию, которая помечет выбранный слайд как предназначенный для `Premium` клиентов. Вот что нужно знать об этом коде:

    - Функция `getSelectedSlideIndex` определяется на следующем шаге. Он возвращает индекс текущего выбранного слайда на основе 1.
    - Значение, возвращаемое `getSelectedSlideIndex` функцией, должно быть удалено, так как метод [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1)) основан на 0.

    ```javascript
    async function addTagToSelectedSlide() {
      await PowerPoint.run(async function(context) {
        let selectedSlideIndex = await getSelectedSlideIndex();
        selectedSlideIndex = selectedSlideIndex - 1;
        const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
        slide.tags.add("CUSTOMER_TYPE", "Premium");
    
        await context.sync();
      });
    }
    ```

2. Следующий код создает метод для получения индекса выбранного слайда. Вот что нужно знать об этом коде:

    - Он использует метод [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) общих API JavaScript.
    - Вызов внедряется `getSelectedDataAsync` в функцию, возвращаемую обещанием. Дополнительные сведения о том, почему и как это сделать, см. в статье "Упаковка общих [API-интерфейсов в](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions) функции, возвращающие обещание".
    - `getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов. В этом сценарии пользователь выбирает только один, поэтому код получает первый (0) слайд, который является единственным выбранным.
    - Значение `index` слайда — это значение на основе 1, которое пользователь видит рядом со слайдом в области эскизов пользовательского интерфейса PowerPoint.

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

3. Следующий код создает функцию для удаления слайдов, помеченных для клиентов уровня "Премиум". Вот что нужно знать об этом коде:

    - Так как `key` свойства `value` тегов `context.sync`будут считываться после них, они должны быть загружены первыми.

    ```javascript
    async function deleteSlidesByAudience() {
      await PowerPoint.run(async function(context) {
        const slides = context.presentation.slides;
        slides.load("tags/key, tags/value");
    
        await context.sync();
    
        for (let i = 0; i < slides.items.length; i++) {
          let currentSlide = slides.items[i];
          for (let j = 0; j < currentSlide.tags.items.length; j++) {
            let currentTag = currentSlide.tags.items[j];
            if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
              currentSlide.delete();
            }
          }
        }
    
        await context.sync();
      });
    }
    ```

## <a name="set-custom-metadata-on-the-presentation"></a>Настройка пользовательских метаданных в презентации

Надстройки также могут применять теги к презентации в целом. Это позволяет использовать теги для метаданных уровня документа, аналогичных тому, как класс [CustomProperty](/javascript/api/word/word.customproperty)используется в Word. Но в отличие от класса Word `CustomProperty` , значение тега PowerPoint может иметь только тип `string`.

Следующий код является примером добавления тега в презентацию. 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

Следующий код является примером удаления тега из презентации. Обратите внимание, что ключ тега передается `delete` методу родительского `TagsCollection` объекта.

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
