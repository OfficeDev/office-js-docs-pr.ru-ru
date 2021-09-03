---
title: Наборы обязательных элементов API ленты
description: Указывает, какие Office платформы и сборки поддерживают динамические API ленты.
ms.date: 05/12/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: aa198009a3d1d16a1c34966516a4ddeee9f7f940
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868738"
---
# <a name="ribbon-api-requirement-sets"></a>Наборы обязательных элементов API ленты

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Набор API ленты поддерживает программный контроль, когда настраиваемые команды надстройки (то есть настраиваемые кнопки ленты и элементы меню) включены и отключены.

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API ленты, Office клиентские приложения, поддерживают этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2013 для Windows<br>(единовременная покупка) | Office 2016 или более поздней Windows<br>(единовременная покупка)   | Office для Windows\*<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac\*<br>(подключено к подписке на Microsoft 365)  | Office в Интернете\*  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| RibbonApi 1.1  | Н/Д | Н/Д | См. поддержку<br>раздел ниже | Н/Д | 16.38 | Ноябрь 2020 г. | Н/Д|
| RibbonApi 1.2  | Н/Д | Н/Д | 2102 (сборка 13801.20294) | Н/Д | скоро | Май 2021 г. | Н/Д|

> **&#42;** API ленты поддерживается только Excel и для этого требуется подписка Microsoft 365.

## <a name="support-for-version-11-on-office-on-windows-subscription"></a>Поддержка версии 1.1 в Office на Windows (подписка)

Версия 1.1 набора требований RibbonApi поддерживается в версии канала 2006 (сборка 13001.20498 или более). Для Office на Windows функция также поддерживается в сборках Semi-Annual channel и Monthly Enterprise Channel, доступных 14 июля 2020 г. или более поздней периодии. Минимальные поддерживаемые сборки для каждого канала:  

|Канал | Версия | Сборка|
|:-----|:-----|:-----|
|Актуальный канал | 2006 или больше | 20266.20266 или более|
|Ежемесячный канал (корпоративный) | 2005 или больше | 12827.20538 или более|
|Ежемесячный канал (корпоративный) | 2004 | 12730.20602 или больше|
|Полугодовой канал (корпоративный) | 2002 или больше | 12527.20880 или больше|

## <a name="more-information"></a>Дополнительные сведения

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Версия и сборка номеров выпусков каналов обновления для Microsoft 365 клиентов](/officeupdates/update-history-microsoft365-apps-by-date)
- [Какая у меня версия Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для Microsoft 365 клиентского приложения](/officeupdates/update-history-microsoft365-apps-by-date)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="ribbon-api-11"></a>API ленты 1.1

API ленты 1.1 — это первая версия API. Сведения об API см. в разделе [Office.ribbon.](/javascript/api/office/office.ribbon)

## <a name="ribbon-api-12"></a>API ленты 1.2

API ленты 1.2 добавляет поддержку контекстных вкладок. Дополнительные сведения см. в статье [Создание пользовательских контекстных вкладок в надстройках Office](../../design/contextual-tabs.md).

> [!NOTE]
> Набор **требований RibbonApi 1.2** еще не поддерживается в манифесте, поэтому не следует указывать его в разделе `<Requirements>` манифест.

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
