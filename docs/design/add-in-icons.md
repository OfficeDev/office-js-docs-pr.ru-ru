---
title: Рекомендации по созданию значков для надстроек Office
description: Общие сведения о проектировании значков, а также о новых и однострочных стилях оформления для команд надстроек.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 35d8e0337b412a9ddebcde5be4db4db802e88269
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093842"
---
# <a name="icons"></a>Значки

Значки — это визуальное представление поведения или концепции. Они часто используются, чтобы отобразить предназначение элементов управления и команд. Визуальные элементы, реалистичные или символические, позволяют выполнять навигацию в пользовательском интерфейсе аналогично значкам в среде пользователя. Эти элементы должны быть простыми, четкими и содержать только необходимые сведения, чтобы пользователи могли быстро проанализировать, какое действие произойдет при выборе элемента управления.

Интерфейсы ленты приложений Office имеют стандартный визуальный стиль. Это поможет обеспечить знакомый пользователям единообразный интерфейс для всех приложений Office. Рекомендации помогут вам создать для своего решения набор ресурсов PNG, полностью совместимых с Office.

Many HTML containers contain controls with iconography. Use Office UI Fabric’s custom font to render Office styled icons in your add-in. Fabric’s icon font contains many glyphs for common Office metaphors that you can scale, color, and style to suit your needs. If you have an existing visual language with your own set of icons, feel free to use it in your HTML canvases. Building continuity with your own brand with a standard set of icons is an important part of any design language. Be careful to avoid creating confusion for customers by conflicting with Office metaphors.

## <a name="design-icons-for-add-in-commands"></a>Создание значков для команд надстроек

С помощью [команд надстроек](add-in-commands.md) можно добавить кнопки, текст и значки в пользовательский интерфейс Office. Значки и метки для кнопок надстроек должны быть понятны и четко определять действие, которое выполняется при выборе пользователем команды. В следующих статьях приводятся рекомендации по оформлению и оформлению, которые помогут вам разрабатывать значки, которые легко интегрируются с Office.

- Линейный стиль Microsoft 365 представлен в разделе стиль с помощью [значка "стильные линии" для надстроек Office](add-in-icons-monoline.md).
- Новые стили Office 2013 +, не связанных с подписками, приведены в статье [новые рекомендации по использованию значков стилей для надстроек Office](add-in-icons-fresh.md).

> [!NOTE]
> Необходимо выбрать один из стилей или другой, и ваша надстройка будет использовать те же значки независимо от того, работает ли он в Microsoft 365 или Office без подписки.

## <a name="see-also"></a>См. также

- [Рекомендации по разработке надстроек](../concepts/add-in-development-best-practices.md)
- [Команды надстроек для Excel, Word и PowerPoint](../design/add-in-commands.md)
