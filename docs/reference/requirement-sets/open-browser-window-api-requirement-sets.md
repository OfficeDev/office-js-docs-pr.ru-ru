---
title: Открытие набора обязательных элементов окна браузера
description: Указывает, какие Office платформы и сборки поддерживают API openBrowserWindow.
ms.date: 09/08/2021
ms.prod: non-product-specific
ms.localizationpriority: medium
ms.openlocfilehash: 1a3518d9efb3b4bf1aec7a9c7713611a130b1c0a
ms.sourcegitcommit: 3fe9e06a52c57532e7968dc007726f448069f48d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/20/2021
ms.locfileid: "59453119"
---
# <a name="open-browser-window-api-requirement-sets"></a>Наборы требований к API окна открытых браузеров

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

Набор API OpenBrowserWindow позволяет надстройке открывать браузер для выполнения задач, которые не всегда можно выполнить в песочнице управления веб-просмотром в самой надстройке; например, скачивание pdf-файла, когда управление веб-просмотром предоставляется Microsoft Edge.

Надстройки Office работают в нескольких версиях Office. В следующей таблице перечислены наборы API API OpenBrowserWindow, Office хост-приложений, поддерживающий этот набор требований, а также номера сборки или версии для Office приложения.

|  Набор обязательных элементов  | Office 2021 или более поздней Windows<br>(единовременная покупка) | Office для Windows<br>(подключено к подписке на Microsoft 365) |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |  Office Online Server  |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| OpenBrowserWindowApi 1.1  | Версия 1810 (сборка 16.0.11001.20074) или более поздней версии | Версия 1810 (сборка 16.0.11001.20074) или более поздней версии | 16.0.0.0 или более поздней | 16.0.0.0 или более поздней | Н/Д | Н/Д|

> [!NOTE]
> Набор требований OpenBrowserWindowApi доступен только следующим образом:
>
> - Excel, PowerPoint, Word: Windows, Mac, iPad
> - Outlook: Windows, Mac

Статьи и разделы с дополнительными сведениями о версиях, номерах сборок и Office Online Server:

- [Версия и сборка номеров выпусков каналов обновления для Приложения Microsoft 365](/officeupdates/update-history-microsoft365-apps-by-date)
- [Какая у меня версия Office](https://support.microsoft.com/office/932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти версию и номер сборки для Office клиентского приложения](/officeupdates/update-history-microsoft365-apps-by-date)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Наборы обязательных элементов общего API для Office

Сведения о наборах обязательных элементов общего API см. в статье [Наборы обязательных элементов общего API для Office](office-add-in-requirement-sets.md).

## <a name="openbrowserwindowapi-11"></a>OpenBrowserWindowApi 1.1

OpenBrowserWindowApi 1.1 — это первая версия API. Подробные сведения об API см. в [справочной Office.context.ui.](/javascript/api/office/office.context#ui)

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
