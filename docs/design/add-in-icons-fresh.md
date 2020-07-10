---
title: Новые рекомендации по значкам стилей для надстроек Office
description: Ознакомьтесь с рекомендациями по использованию новых значков значков стилей в надстройках Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 7f29de70712448e9ee7458db864fb40746412153
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093933"
---
# <a name="fresh-style-icon-guidelines-for-office-add-ins"></a>Новые рекомендации по значкам стилей для надстроек Office

В версиях Office для Office 2013 + (не подписке) используется значки в стиле Microsoft Office. Если вы предпочитаете, чтобы значки выглядели как однострочный стиль Microsoft 365, ознакомьтесь с [разделом стиль стильных значков для надстроек Office](add-in-icons-monoline.md).

## <a name="office-fresh-visual-style"></a>Новый визуальный стиль Office

Новые значки включают только важные элементы коммуникативе. Вспомогательные элементы, включая перспективу, градиенты и источник света, удалены. Упрощение значков способствует ускорению анализа команд и элементов управления. Используйте этот стиль для оптимального размещения клиентов, не использующих подписки на Office.

## <a name="best-practices"></a>Рекомендации

При создании значков следуйте перечисленным ниже рекомендациям.

|Правильно|Неправильно|
|:---|:---|
|Сохраняйте визуальные элементы простыми и понятными, чтобы сосредоточиться на ключевых элементах общения.| Не используйте артефакты, которые визуально искажают изображение значка.|
|Используйте язык значков Office для представления поведения или концепций.|Не изменяйте глифы Office UI Fabric для команд надстройки на ленте приложений Office или контекстных меню. Значки Fabric отличаются по стилю, поэтому не будут совпадать.|
|Повторно используйте общие визуальные метафоры Office, например кисть для форматирования или увеличительное стекло для поиска.|Не используйте повторно визуальные метафоры для различных команд. Добавление одинаковых значков для различных действий и концепций может привести к путанице. |
|Перерисуйте свои значки, чтобы уменьшить или увеличить их. Перерисуйте контуры, углы и скругленные края, чтобы повысить четкость линий. |Не изменяйте размеры значков, сжимая или растягивая их. Это может привести к ухудшению визуального качества и непонятному изображению действий. Сложные значки, созданные в большем размере, могут потерять четкость при их уменьшении без перерисовки. |
|Use a white fill for accessibility. Most objects in your icons will require a white background to be legible across Office UI themes and in high-contrast modes.  |Ваш логотип или торговая марка могут и не указывать на функции определенной команды надстройки. Торговые знаки не всегда можно легко узнать, если они обозначены значками меньшего размера, а также когда к ним применены модификаторы. Метки торговых марок часто конфликтуют с стилями значков ленты приложения Office и могут конкурировать за пользовательское вмешательство в насыщенной среде. |
|Используйте формат PNG с прозрачным фоном. ||
|Избегайте использования в значках локализуемого содержимого, а также типографских символов, индикаторов абзацев без выравнивания и вопросительных знаков. ||

## <a name="icon-size-recommendations-and-requirements"></a>Рекомендации и требования, применяющиеся к размерам значков

Значки Office на рабочем столе являются растровыми изображениями. Различные размеры будут отображаться в зависимости от установленного пользователем разрешения экрана и сенсорного режима. Используйте все восемь поддерживаемых размеров, чтобы обеспечить лучшее представление для всех поддерживаемых разрешений и контекстов. Ниже перечислены поддерживаемые размеры, из которых обязательными являются три:

- 16 пк (обязательный);
- 20 пк;
- 24 пк;
- 32 пк (обязательный);
- 40 пк;
- 48 пк;
- 64 пк (рекомендуется, лучший вариант для компьютера Mac);
- 80 пк (обязательный).

Не сжимайте значки, а перерисуйте их для каждого размера.

![Рисунок, на котором показана рекомендация не сжимать значки, а изменить их размер](../images/icon-resizing.png)

<!--
The following table shows the icon sizes that render for different modes at different DPI settings.

|DPI |**Small**||**Medium**||**Large**||**Extra large**|
|:---|:---|:---|:---|:---|:---|:---|:---|
|    |**Mouse**|**Touch**|**Mouse**|**Touch**|**Mouse**|**Touch**|-|
|100%|16px|20px|24px||32px|40px|48px|
|125%|20px|24px|||40px|48px|60px|
|150%|24px|24px|36px||48px|48px|72px|
|200%|32px|40px|48px||64px|80px|96px|
|250%|40px||||80px||120px|
|300%|48px||||96px||144px

> [!NOTE]
> At DPI settings of 150% or greater, the icon does not get swapped out for a larger size when Touch mode is engaged. At DPI settings greater than 250%, Touch mode is turned off by default.

The following table lists the locations for certain icon sizes.

|Location|100% DPI|200% DPI|250% DPI|
|:-------|:-------|:-------|:-------|
|Small ribbon button|16px|32px|40px|
|Contextual menu|16px|32px|40px|
|Quick access toolbar (QAT)|16px|32px|40px|
|Large ribbon icon|32px|64px|80px|

-->

## <a name="icon-anatomy-and-layout"></a>Структура и схема значка

Office icons are typically comprised of a base element with action and conceptual modifiers overlayed. Action modifiers represent concepts such as add, open, new, or close. Conceptual modifiers represent status, alteration, or a description of the icon.

To create commands that align with the Office UI, follow layout guidelines for the base element and modifiers. This ensures that your commands look professional and that your customers will trust your add-in. If you make exceptions to these guidelines, do so intentionally.

На следующем изображении показана схема расположения базовых элементов и модификаторов для значка Office.

![Изображение, на котором базовый элемент значка показан в центре, модификатор действия — в левом верхнем углу, а другой модификатор — в нижнем правом углу](../images/icon-layouts.png)

- Размещайте базовые элементы в центре пиксельной рамки, оставляя немного места по краям.
- Модификаторы действия располагайте в верхнем левом углу.
- Концептуальные модификаторы размещайте в нижнем правом углу.
- Limit the number of elements in your icons. At 32px, limit the number of modifiers to a maximum of two. At 16px, limit the number of modifiers to one.

### <a name="base-element-padding"></a>Отступ вокруг базового элемента

Размещайте базовые элементы единообразно для всех размеров. Если у вас не получается разместить базовые элементы в центре рамки, расположите их в левом верхнем углу, оставив несколько дополнительных пикселей в правом нижнем углу. Для достижения лучших результатов примените рекомендации по заполнению, приведенные в таблице в следующем разделе.

### <a name="modifiers"></a>Модификаторы

All modifiers should have a 1px transparent cutout between each element, including the background. Elements should not directly overlap. Create whitespace between rules and edges. Modifiers can vary slightly in size, but use these dimensions as a starting point.

|**Размер значка**|**Отступ вокруг базового элемента**|**Размер модификатора**|
|:---|:---|:---|
|16 пк|нуль|9 пк|
|20 пк|1 пк|10 пк|
|24 пк|1 пк|12 пк|
|32 пк|2 пк|14 пк|
|40 пк|2 пк|20 пк|
|48 пк|3 пк|22 пк|
|64 пк|5 пк|29 пк|
|80 пк|5 пк|38 пк|

## <a name="icon-colors"></a>Цвета значков

> [!NOTE]
> Эти руководства по цветам применяются к значкам ленты, используемым в [командах надстроек](add-in-commands.md). Эти значки не отрисовываются с помощью Microsoft UI Fabric, и цветовая палитра отличается от палитры, описанной на странице [Microsoft UI Fabric | Colors | Shared](https://fluentfabric.azurewebsites.net/#/color/shared).

Office icons have a limited color palette. Use the colors listed in the following table to guarantee seamless integration with the Office UI. Apply the following guidelines to the use of color:

- Use color to communicate meaning rather than for embellishment. It should highlight or emphasize an action, status, or an element that explicitly differentiates the mark. 
- If possible, use only one additional color beyond gray. Limit additional colors to two at the most.
- Colors should have a consistent appearance in all icon sizes. Office icons have slightly different color palettes for different icon sizes. 16px and smaller icons are slightly darker and more vibrant than 32px and larger icons. Without these subtle adjustments, colors appear to vary across sizes.

|**Название цвета**|**RGB**|**Шестнадцатеричный код**|**Цвет**|**Категория**|
|:---|:---|:---|:---|:---|
|Серый цвет текста (80)|80, 80, 80|#505050| ![Изображение серого цвета текста (80)](../images/color-text-gray-80.png) |Текст|
|Серый цвет текста (95)|95, 95, 95|#5F5F5F| ![Изображение серого цвета текста (95)](../images/color-text-gray-95.png) |Текст|
|Серый цвет текста (105)|105, 105, 105|#696969| ![Изображение серого цвета текста (105)](../images/color-text-gray-105.png) |Текст|
|Темно-серый (32)|128, 128, 128|#808080| ![Изображение темно-серого цвета (32)](../images/color-dark-gray-32.png) |32 и больше|
|умеренно серый (32)|158, 158, 158|#9E9E9E| ![Изображение умеренно серого цвета (32)](../images/color-medium-gray-32.png) |32 и больше|
|Светло-серый (ВСЕ)|179, 179, 179|#B3B3B3| ![Изображение светло-серого цвета (для всех размеров)](../images/color-light-gray-all.png) |Все размеры|
|Темно-серый (16)|114, 114, 114|#727272| ![Изображение темно-серого цвета (16)](../images/color-dark-gray-16.png) |16 и меньше|
|Умеренно серый (16)|144, 144, 144|#909090| ![Изображение умеренно серого цвета (16)](../images/color-medium-gray-16.png) |16 и меньше|
|Синий (32)|77, 130, 184|#4d82B8| ![Изображение синего цвета (32)](../images/color-blue-32.png) |32 и больше|
|Синий (16)|74, 125, 177|#4A7DB1| ![Изображение синего цвета (16)](../images/color-blue-16.png) |16 и меньше|
|Желтый (ВСЕ)|234, 194, 130|#EAC282| ![Изображение желтого цвета для всех размеров](../images/color-yellow-all.png) |Все размеры|
|Оранжевый (32)|231, 142, 70|#E78E46| ![Изображение оранжевого цвета (32)](../images/color-orange-32.png) |32 и больше|
|Оранжевый (16)|227, 142, 70|#E3751C| ![Изображение оранжевого цвета (16)](../images/color-orange-16.png) |16 и меньше|
|Розовый (ВСЕ)|230, 132, 151|#E68497| ![Изображение розового цвета для всех размеров](../images/color-pink-all.png) |Все размеры|
|Зеленый (32)|118, 167, 151|#76A797| ![Изображение зеленого цвета (32)](../images/color-green-32.png) |32 и больше|
|Зеленый (16)|104, 164, 144|#68A490| ![Изображение зеленого цвета (16)](../images/color-green-16.png) |16 и меньше|
|Красный (32)|216, 99, 68|#D86344| ![Изображение красного цвета (32)](../images/color-red-32.png) |32 и больше|
|Красный (16)|214, 85, 50|#D65532| ![Изображение красного цвета (16)](../images/color-red-16.png) |16 и меньше|
|Сиреневый (32)|152, 104, 185|#9868B9| ![Изображение сиреневого цвета (32)](../images/color-purple-32.png) |32 и больше|
|Сиреневый (16)|137, 89, 171|#8959AB| ![Изображение сиреневого цвета (16)](../images/color-purple-16.png) |16 и меньше|

## <a name="icons-in-high-contrast-modes"></a>Значки в режимах высокой контрастности

Office icons are designed to render well in high contrast modes. Foreground elements are well differentiated from backgrounds to maximize legibility and enable recoloring. In high contrast modes, Office will recolor any pixel of your icon with a red, green, or blue value less than 190 to full black. All other pixels will be white. In other words, each RGB channel is assessed where 0-189 values are black and 190-255 values are white. Other high-contrast themes recolor using the same 190 value threshold but with different rules. For example, the high-contrast white theme will recolor all pixels greater than 190 opaque but all other pixels as transparent. Apply the following guidelines to maximize legibility in high-contrast settings:

- Старайтесь разграничивать элементы переднего и заднего планов с учетом порогового значения 190.
- Следуйте стилям оформления значков Office.
- Используйте для значков цвета из нашей палитры.
- Не рекомендуется использовать градиенты.
- Избегайте больших блоков цветов с похожими значениями.
