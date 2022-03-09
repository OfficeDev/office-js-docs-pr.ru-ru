---
title: Null значения в надстройки Word
description: Узнайте, как работать со значениями null в надстройке Word.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e21677dafcaaaa7e9e9164ef18c82f49820298d6
ms.sourcegitcommit: 9d930b4c77c342246607aef30479e31fdbdd47f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63353861"
---
# <a name="null-values-in-word-add-ins"></a>Null значения в надстройки Word

`null` имеет особые последствия в API Word JavaScript. Он используется для представления значений по умолчанию или без форматирования.

## <a name="null-property-values-in-the-response"></a>Значения свойств null в ответе

Форматирование свойств, таких как [цвет](/javascript/api/word/word.font#word-word-font-color-member) , `null` будет содержать значения в ответе при существовании различных значений в указанном [диапазоне](/javascript/api/word/word.range). Например, если вы получаете диапазон и загружаете его свойство `range.font.color`:

- Если весь текст в диапазоне имеет одинаковый цвет шрифта, `range.font.color` указывает этот цвет.
- Если в диапазоне используется несколько цветов шрифтов, свойство `range.font.color` имеет значение `null`.
