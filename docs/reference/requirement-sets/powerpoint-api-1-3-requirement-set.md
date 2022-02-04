---
title: PowerPoint API JavaScript установлено 1.3
description: Сведения о наборе требований PowerPointApi 1.3.
ms.date: 12/14/2021
ms.prod: powerpoint
ms.localizationpriority: medium
---

# <a name="whats-new-in-powerpoint-javascript-api-13"></a>Новые возможности в PowerPoint API JavaScript 1.3

PowerPointApi 1.3 добавила дополнительную поддержку для управления слайдами и настраиваемой метки.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Управление слайдами](../../powerpoint/add-slides.md) | Добавляет поддержку для добавления слайдов, а также управления макетами слайдов и мастерами слайдов. | [Slide](/javascript/api/powerpoint/powerpoint.slide)<br>[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)<br>[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|
| [Tags](../../powerpoint/tagging-presentations-slides-shapes.md) | Позволяет надстройкам прикреплять настраиваемые метаданные в виде пар с ключевым значением. | [Tag](/javascript/api/powerpoint/powerpoint.tag) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены PowerPoint API JavaScript, установленный 1.3. Полный список всех API PowerPoint JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. PowerPoint [API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[AddSlideOptions](/javascript/api/powerpoint/powerpoint.addslideoptions)|[layoutId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-layoutid-member)|Указывает ID макета слайда, который будет использоваться для нового слайда.|
||[slideMasterId](/javascript/api/powerpoint/powerpoint.addslideoptions#powerpoint-powerpoint-addslideoptions-slidemasterid-member)|Указывает ID мастера слайдов, который будет использоваться для нового слайда.|
|[Presentation](/javascript/api/powerpoint/powerpoint.presentation)|[slideMasters](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-slidemasters-member)|Возвращает коллекцию объектов `SlideMaster` , которые находятся в презентации.|
||[теги](/javascript/api/powerpoint/powerpoint.presentation#powerpoint-powerpoint-presentation-tags-member)|Возвращает коллекцию тегов, присоединенных к презентации.|
|[Shape](/javascript/api/powerpoint/powerpoint.shape)|[delete()](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-delete-member(1))|Удаляет фигуру из коллекции фигур.|
||[id](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-id-member)|Получает уникальный ID формы.|
||[теги](/javascript/api/powerpoint/powerpoint.shape#powerpoint-powerpoint-shape-tags-member)|Возвращает коллекцию тегов в форме.|
|[ShapeCollection](/javascript/api/powerpoint/powerpoint.shapecollection)|[getCount()](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getcount-member(1))|Получает количество фигур в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitem-member(1))|Получает форму с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemat-member(1))|Получает фигуру с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-getitemornullobject-member(1))|Получает форму с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.shapecollection#powerpoint-powerpoint-shapecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Slide](/javascript/api/powerpoint/powerpoint.slide)|[макет](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-layout-member)|Получает макет слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-shapes-member)|Возвращает коллекцию фигур на слайде.|
||[slideMaster](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-slidemaster-member)|Получает объект `SlideMaster` , который представляет содержимое слайда по умолчанию.|
||[теги](/javascript/api/powerpoint/powerpoint.slide#powerpoint-powerpoint-slide-tags-member)|Возвращает коллекцию тегов на слайде.|
|[SlideCollection](/javascript/api/powerpoint/powerpoint.slidecollection)|[add(options?: PowerPoint. AddSlideOptions)](/javascript/api/powerpoint/powerpoint.slidecollection#powerpoint-powerpoint-slidecollection-add-member(1))|Добавляет новый слайд в конце коллекции.|
|[SlideLayout](/javascript/api/powerpoint/powerpoint.slidelayout)|[id](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-id-member)|Получает уникальный ID макета слайда.|
||[name](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-name-member)|Получает имя макета слайда.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidelayout#powerpoint-powerpoint-slidelayout-shapes-member)|Возвращает коллекцию фигур в макете слайда.|
|[SlideLayoutCollection](/javascript/api/powerpoint/powerpoint.slidelayoutcollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getcount-member(1))|Получает количество макетов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitem-member(1))|Получает макет с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemat-member(1))|Получает макет с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-getitemornullobject-member(1))|Получает макет с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidelayoutcollection#powerpoint-powerpoint-slidelayoutcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlideMaster](/javascript/api/powerpoint/powerpoint.slidemaster)|[id](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-id-member)|Получает уникальный ID мастера слайдов.|
||[макеты](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-layouts-member)|Получает коллекцию макетов, предоставленных мастером слайдов для слайдов.|
||[name](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-name-member)|Получает уникальное имя мастера слайдов.|
||[shapes](/javascript/api/powerpoint/powerpoint.slidemaster#powerpoint-powerpoint-slidemaster-shapes-member)|Возвращает коллекцию фигур в мастере слайдов.|
|[SlideMasterCollection](/javascript/api/powerpoint/powerpoint.slidemastercollection)|[getCount()](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getcount-member(1))|Получает число мастеров слайдов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitem-member(1))|Получает мастер слайдов с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemat-member(1))|Получает мастер слайдов с помощью нулевого индекса в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-getitemornullobject-member(1))|Получает мастер слайдов с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.slidemastercollection#powerpoint-powerpoint-slidemastercollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Tag](/javascript/api/powerpoint/powerpoint.tag)|[key](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-key-member)|Получает уникальный ID тега.|
||[value](/javascript/api/powerpoint/powerpoint.tag#powerpoint-powerpoint-tag-value-member)|Получает значение тега.|
|[TagCollection](/javascript/api/powerpoint/powerpoint.tagcollection)|[add(key: string, value: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-add-member(1))|Добавляет новый тег в конце коллекции.|
||[delete (key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-delete-member(1))|Удаляет тег с заданным в `key` этой коллекции.|
||[getCount()](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getcount-member(1))|Получает количество тегов в коллекции.|
||[getItem(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitem-member(1))|Получает тег с помощью уникального ID.|
||[getItemAt(index: number)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemat-member(1))|Получает тег с использованием нулевого индекса в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-getitemornullobject-member(1))|Получает тег с помощью уникального ID.|
||[items](/javascript/api/powerpoint/powerpoint.tagcollection#powerpoint-powerpoint-tagcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [PowerPoint справочная документация по API JavaScript](/javascript/api/powerpoint?view=powerpoint-js-1.3&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для PowerPoint](powerpoint-api-requirement-sets.md)
