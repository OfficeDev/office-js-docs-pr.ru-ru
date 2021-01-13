---
title: Наборы обязательных элементов API ленты
description: Указывает, какие платформы и сборки Office поддерживают динамические API ленты.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839987"
---
# <a name="ribbon-api-requirement-sets"></a>Наборы обязательных элементов API ленты

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Набор API ленты поддерживает программный контроль того, когда настраиваемые команды надстройки (то есть настраиваемые кнопки ленты и элементы меню) включены и отключены.

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований API ленты, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборки или версии приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 или более поздней версии для Windows<br>(единовременная покупка)   | Office для Windows\*<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac\*<br>(подключено к подписке на Microsoft 365)  | Office в Интернете\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | Н/Д | Н/Д | См. службу поддержки<br>раздел ниже | Недоступно | 16.38 | Ноябрь 2020 г. | Недоступно|

> **&#42;** API ленты поддерживается только в Excel и требует подписки на Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Поддержка Office для Windows (подписка)

Набор требований поддерживается в канале Consumer Channel версии 2006 (сборка 13001.20498 или более новой). Для Office для Windows эта функция также поддерживается в сборках канала Semi-Annual Channel и Monthly Enterprise Channel, доступных 14 июля 2020 г. или более поздней версии. Минимальные поддерживаемые сборки для каждого канала:  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2006 или более | 20266.20266 или более|
|Ежемесячный канал (корпоративный) | 2005 или более | 12827.20538 или более|
|Ежемесячный канал (корпоративный) | 2004 | 12730.20602 или более|
|Полугодовой канал (корпоративный) | 2002 или более | 12527.20880 или более|

## <a name="more-information"></a>Дополнительная информация

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и выпусков каналов обновления для клиентов Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для клиентского приложения Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> Набор **требований RibbonApi 1.1** еще не поддерживается в манифесте, поэтому его нельзя указать в разделе `<Requirements>` манифеста.


## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API ленты 1.1

API ленты 1.1 является первой версией API. Подробные сведения об API см. в справочном разделе [office.ribbon.](/javascript/api/office/office.ribbon)

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)