---
title: Наборы обязательных элементов API ленты
description: Указывает, какие платформы и сборки Office поддерживают динамические API ленты.
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 878670367b253fa7700434681244b43b9cfa36a7
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996517"
---
# <a name="ribbon-api-requirement-sets"></a>Наборы обязательных элементов API ленты

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Набор API ленты поддерживает программное управление при включении и отключении пользовательских команд надстроек (то есть кнопок ленты и их элементов меню).

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы требований API ленты, клиентские приложения Office, которые поддерживают этот набор требований, а также номера сборок или версий для приложения Office.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 или более поздней версии в Windows<br>(единовременная покупка)   | Office для Windows\*<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac\*<br>(подключено к подписке на Microsoft 365)  | Office в Интернете\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| Риббонапи 1,1  | Н/Д | Н/Д | Ознакомьтесь со статьей поддержка<br>раздел ниже | Н/Д | 16,38 | Ноябрь, 2020 | Н/Д|

> **&#42;** API ленты поддерживается только в Excel и требует подписки Microsoft 365.

## <a name="office-on-windows-subscription-support"></a>Поддержка Office в Windows (подписка)

Набор обязательных элементов поддерживается в канале потребителей версии 2006 (сборка, 13001,20498 или выше). Для Office в Windows эта функция также поддерживается в сборках каналов Semi-Annual и месячных корпоративных каналах, доступных в июле 14th 2020 или более поздних версий. Ниже приведены минимальные поддерживаемые сборки для каждого канала.  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2006 или выше | 20266,20266 или выше|
|Ежемесячный канал (корпоративный) | 2005 или выше | 12827,20538 или выше|
|Ежемесячный канал (корпоративный) | 2004 | 12730,20602 или выше|
|Полугодовой канал (корпоративный) | 2002 или выше | 12527,20880 или выше|

## <a name="more-information"></a>Дополнительные сведения

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Номера версий и сборок выпусков из канала обновления для клиентов Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для клиентского приложения Microsoft 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> Набор требований **риббонапи 1,1** пока не поддерживается в манифесте, поэтому его нельзя указать в разделе манифеста `<Requirements>` .


## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API ленты 1,1

API ленты 1,1 — это первая версия API. Более подробную информацию об API можно узнать в разделе [Office. Ribbon ](/javascript/api/office/office.ribbon) .

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
