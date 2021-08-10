---
title: Используйте настраиваемые теги на презентациях, слайдах и фигурах в PowerPoint
description: Узнайте, как использовать теги для настраиваемой метаданных о презентациях, слайдах и фигурах.
ms.date: 04/08/2021
localization_priority: Normal
ms.openlocfilehash: 9ae86906a2ac69cb79adac34fa4e923a9bc218a7dc8a7e5bdefd63300b589da5
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093660"
---
# <a name="use-custom-tags-for-presentations-slides-and-shapes-in-powerpoint"></a>Используйте настраиваемые теги для презентаций, слайдов и фигур в PowerPoint

Надстройка может прикреплять настраиваемые метаданные в виде пар значений ключей, называемых "тегами", к презентациям, определенным слайдам и определенным фигурам на слайде.

> [!IMPORTANT]
> API для тегов находятся в предварительном просмотре. Поэкспериментируйте с ними в среде разработки или тестирования, но не добавляйте их в производственную надстройка.

Существует два основных сценария использования тегов:

- При применении к слайду или фигуре тег позволяет классифицовать объект для пакетной обработки. Например, предположим, что в презентации есть слайды, которые следует включить в презентации восточного региона, но не западного региона. Кроме того, существуют альтернативные слайды, которые должны показываться только на Западе. Надстройка может создать тег с ключом и значением и применить его к слайдам, которые следует `REGION` `East` использовать только на Востоке. Значение тега заказано для слайдов, которые должны показываться только в `West` западном регионе. Перед презентацией на Востоке кнопка в коде надстройки выполняет циклы через все слайды, проверяя значение `REGION` тега. Слайды, в которых `West` область удалена. Затем пользователь закрывает надстройку и запускает слайд-шоу.
- При применении к презентации тег фактически является настраиваемой свойством в документе презентации (аналогично [CustomProperty](/javascript/api/word/word.customproperty) в Word).

## <a name="tag-slides-and-shapes"></a>Слайды и фигуры тегов

Тег — это пара значений ключа, где значение всегда типа и представлено `string` объектом [Tag.](/javascript/api/powerpoint/powerpoint.tag) Каждый тип родительского объекта, например [объект Presentation,](/javascript/api/powerpoint/powerpoint.presentation) [Slide](/javascript/api/powerpoint/powerpoint.slide)или [Shape,](/javascript/api/powerpoint/powerpoint.shape) имеет свойство `tags` типа [TagsCollection.](/javascript/api/powerpoint/powerpoint.tagcollection)

### <a name="add-update-and-delete-tags"></a>Добавление, обновление и удаление тегов

Чтобы добавить тег к объекту, вызовите [метод TagCollection.add](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_) свойства родительского `tags` объекта. Следующий код добавляет два тега к первому слайду презентации. Вот что нужно знать об этом коде:

- Первым параметром метода `add` является ключ в паре значение ключа. 
- Второй параметр — это значение.
- Ключ находится в верхних буквах. Это не является строго обязательным для метода; однако ключ всегда хранится в PowerPoint как верхний шкаф, и некоторые методы, связанные с тегами, требуют, чтобы ключ был выражен в верхнем шкафу, поэтому мы рекомендуем в качестве рекомендации использовать верхний шкаф в коде для `add` ключа тега. 

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

Метод `add` также используется для обновления тега. Следующий код изменяет значение `PLANET` тега.

```javascript
async function updateTag() {
  await PowerPoint.run(async function(context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}
```

Чтобы удалить тег, позвоните методу на родительском объекте и передайте ключ тега `delete` `TagsCollection` в качестве параметра. Например, см. [в примере Set custom metadata on the presentation.](#set-custom-metadata-on-the-presentation)

### <a name="use-tags-to-selectively-process-slides-and-shapes"></a>Использование тегов для выборочной обработки слайдов и фигур

Рассмотрим следующий сценарий: Contoso Consulting имеет презентацию, которая будет показываться всем новым клиентам. Но некоторые слайды должны показываться только тем клиентам, которые заплатили за состояние "премиум". Перед показом презентации для клиентов, не взмываюых к премиум-классам, они делают ее копию и удаляют слайды, которые должны видеть только клиенты премиум-класса. Надстройка позволяет Contoso теги, какие слайды для премиум-клиентов и удалить эти слайды при необходимости. В следующем списке описаны основные этапы кодирования для создания этой функции.

1. Создайте метод, который помечет выбранный в настоящее время слайд как предназначенный для `Premium` клиентов. Вот что нужно знать об этом коде:

    - Функция `getSelectedSlideIndex` определяется на следующем шаге. Он возвращает индекс на основе 1 выбранного слайда.
    - Значение, возвращаемого функцией, должно быть отсоединимо, так как метод `getSelectedSlideIndex` [SlideCollection.getItemAt](/javascript/api/powerpoint/powerpoint.slidecollection#getItemAt_index_) основан на 0.

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

2. Следующий код создает метод получения индекса выбранного слайда. Вот что нужно знать об этом коде:

    - Он использует метод [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__callback_) общих API JavaScript.
    - Вызов встроен в функцию возврата `getSelectedDataAsync` обещаний. Дополнительные сведения о том, почему и как это сделать, см. в этой ссылке [Wrap Common API в функциях возврата обещаний.](../develop/asynchronous-programming-in-office-add-ins.md#wrap-common-apis-in-promise-returning-functions)
    - `getSelectedDataAsync` возвращает массив, так как можно выбрать несколько слайдов. В этом сценарии пользователь выбрал только один, поэтому код получает первый (0-й) слайд, который является единственным выбранным.
    - Значение слайда — это 1-базирующееся значение, что пользователь видит рядом со слайдом в области PowerPoint `index` интерфейса пользовательского интерфейса.

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

3. В следующем коде создается метод удаления слайдов, помеченных для премиум-клиентов. Вот что нужно знать об этом коде:

    - Так как свойства и свойства тегов будут читаться после загрузки, они должны `key` `value` быть `context.sync` загружены в первую очередь.

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

## <a name="set-custom-metadata-on-the-presentation"></a>Настройка настраиваемой метаданных на презентации

Надстройки также могут применять теги к презентации в целом. Это позволяет использовать теги для метаданных на уровне документов, аналогичные использованию класса [CustomProperty](/javascript/api/word/word.customproperty)в Word. Но в отличие от класса Word, значение тега PowerPoint `CustomProperty` может быть только типа `string` .

Следующий код — пример добавления тега в презентацию. 

```javascript
async function addPresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.add("SECURITY", "Internal-Audience-Only");

    await context.sync();
  });
}
```

Следующий код — пример удаления тега из презентации. Обратите внимание, что ключ тега передается `delete` методу родительского `TagsCollection` объекта.

```javascript
async function deletePresentationTag() {
  await PowerPoint.run(async function (context) {
    let presentationTags = context.presentation.tags;
    presentationTags.delete("SECURITY");

    await context.sync();
  });
}
```
