---
title: PowerPoint API JavaScript установлено 1.3
description: Сведения о наборе требований PowerPointApi 1.3.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
ms.openlocfilehash: 74f17923f7bc8b26416c39bdbbeea9cc0a94029a
ms.sourcegitcommit: e44a8109d9323aea42ace643e11717fb49f40baa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/15/2021
ms.locfileid: "61514249"
---
# <a name="whats-new-in-powerpoint-javascript-api-13"></a>Новые возможности в PowerPoint API JavaScript 1.3

PowerPointApi 1.3 добавила дополнительную поддержку для управления слайдами и настраиваемой метки.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Управление слайдами](../../powerpoint/add-slides.md) | Добавляет поддержку для добавления слайдов, а также управления макетами слайдов и мастерами слайдов. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | Позволяет надстройкам прикреплять настраиваемые метаданные в виде пар с ключевым значением. | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, установленный 1.3. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. PowerPoint [API JavaScript.](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#layoutId)|Указывает ID макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#slideMasterId)|Указывает ID мастера слайдов, который будет использоваться для нового слайда.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#slideMasters)|Возвращает коллекцию `SlideMaster` объектов, которые находятся в презентации.|
||[tags](/javascript/api/powerpoint/powerpoint.presentation#tags)|Возвращает коллекцию тегов, присоединенных к презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#delete__)|Удаляет фигуру из коллекции фигур.|
||[id](/javascript/api/powerpoint/powerpoint.shape#id)|Получает уникальный ID формы.|
||[tags](/javascript/api/powerpoint/powerpoint.shape#tags)|Возвращает коллекцию тегов в форме.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#getCount__)|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#getItem_key_)|Получает форму с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemAt_index_)|Получает фигуру с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#getItemOrNullObject_id_)|Получает форму с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[макет](/javascript/api/powerpoint/powerpoint.slide#layout)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#shapes)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#slideMaster)|Получает `SlideMaster` объект, который представляет содержимое слайда по умолчанию.|
||[tags](/javascript/api/powerpoint/powerpoint.slide#tags)|Возвращает коллекцию тегов на слайде.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#add_options_)|Добавляет новый слайд в конце коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#id)|Получает уникальный ID макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#name)|Получает имя макета слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#shapes)|Возвращает коллекцию фигур в макете слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getCount__)|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItem_key_)|Получает макет с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemAt_index_)|Получает макет с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#getItemOrNullObject_id_)|Получает макет с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#id)|Получает уникальный ID мастера слайдов.|
||[макеты](/javascript/api/powerpoint/powerpoint.slidemaster#layouts)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#name)|Получает уникальное имя мастера слайдов.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#shapes)|Возвращает коллекцию фигур в мастере слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#getCount__)|Получает число мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItem_key_)|Получает мастер слайдов с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemAt_index_)|Получает мастер слайдов с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#getItemOrNullObject_id_)|Получает мастер слайдов с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#key)|Получает уникальный ID тега.|
||[value](/javascript/api/powerpoint/powerpoint.tag#value)|Получает значение тега.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#add_key__value_)|Добавляет новый тег в конце коллекции.|
||[delete (key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#delete_key_)|Удаляет тег с заданным `key` в этой коллекции.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#getCount__)|Получает количество тегов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItem_key_)|Получает тег с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemAt_index_)|Получает тег с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#getItemOrNullObject_key_)|Получает тег с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [PowerPoint справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
