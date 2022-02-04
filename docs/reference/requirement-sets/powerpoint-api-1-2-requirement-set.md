---
title: PowerPoint API JavaScript установлено 1.2
description: Сведения о наборе требований PowerPointApi 1.2.
ms.date: 01/27/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-12"></a>Новые возможности в PowerPoint API JavaScript 1.2

PowerPointApi 1.2 добавила поддержку для вставки слайдов из другой презентации в текущую презентацию и удаления слайдов.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Вставка и удаление слайдов](../../powerpoint/insert-slides-into-presentation.md) | Позволяет вставлять существующие слайды в текущую презентацию из другой презентации, а также возможность удаления слайдов. | [Slide.delete](/javascript/api/powerpoint/powerpoint.slide#delete--), [Presentation.insertSlidesFromBase64](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, установленный 1.2. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. PowerPoint [API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[InsertSlideOptions](/javascript/api/powerpoint/powerpoint.insertslideoptions)|[форматирование](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-formatting-member)|Указывает форматирование, которое необходимо использовать во время вставки слайда.|
||[sourceSlideIds](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-sourceslideids-member)|Указывает слайды из исходных презентаций, которые будут вставлены в текущую презентацию.|
||[targetSlideId](/javascript/api/powerpoint/powerpoint.insertslideoptions#powerpoint-powerpoint-insertslideoptions-targetslideid-member)|Указывает, где в презентации будут вставлены новые слайды.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[insertSlidesFromBase64(base64File: string, options?: PowerPoint. InsertSlideOptions)](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-insertslidesfrombase64-member(1))|Вставляет указанные слайды из презентации в текущую презентацию.|
||[слайды](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slides-member)|Возвращает упорядоченную коллекцию слайдов в презентации.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[delete()](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-delete-member(1))|Удаляет слайд из презентации.|
||[id](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-id-member)|Получает уникальный ID слайда.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getcount-member(1))|Получает количество слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitem-member(1))|Получает слайд с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemat-member(1))|Получает слайд с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-getitemornullobject-member(1))|Получает слайд с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [PowerPoint справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.2&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
