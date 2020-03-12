---
title: Наборы обязательных элементов API JavaScript для Excel
description: Сведения о наборе обязательных элементов надстройки Office для сборок Excel
ms.date: 03/11/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: b6e1570d7487e552197201d12f9a783f18a30fe3
ms.sourcegitcommit: 05b73cdec5f4db7f0b8d48a5a552ee296a0332ca
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/11/2020
ms.locfileid: "42600706"
---
# <a name="excel-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Excel

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Доступность набора обязательных элементов

Надстройки Excel работают в нескольких версиях Office, включая Office 2016 или более поздние версии для Windows, а также Office в Интернете, Office для Mac и Office для iPad. В приведенной ниже таблице перечислены наборы обязательных элементов для Excel, ведущие приложения Office, которые поддерживают все наборы обязательных инструментов, а также номера сборок или версий для этих приложений.

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов или `ExcelApiOnline`, следует ссылаться на **рабочую** библиотеку в сети CDN https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Сведения об использовании API предварительных версий см. в статье [Предварительные версии API JavaScript для Excel](excel-preview-apis.md).

|  Набор обязательных элементов  |  Office для Windows<br>(версия, подключенная к подписке на Office 365)  |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Предварительная версия](excel-preview-apis.md)  | Применяйте последнюю версию Office для использования предварительных версий API (может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider)) |
| [ExcelApiOnline](excel-api-online-requirement-set.md) | Н/Д | Н/Д | Н/Д | Последние (см. [набор обязательных элементов, стр.](./excel-api-online-requirement-set.md)) |
| [ExcelApi 1.10](excel-api-1-10-requirement-set.md) | Версия 1907 (сборка 11929.20306) или более поздняя | 2.30 или более поздняя версия | 16.30 или более поздняя версия | Октябрь 2019 г. |
| [ExcelApi 1.9](excel-api-1-9-requirement-set.md)  | Версия 1903 (сборка 11425.20204) или более поздняя | 2.24 или более поздняя версия | 16.24 или более поздняя версия | Май 2019 г. |
| [ExcelApi 1.8](excel-api-1-8-requirement-set.md)  | Версия 1808 (сборка 10730.20102) или более поздняя | 2.17 или более поздняя | 16.17 или более поздняя | Сентябрь 2018 г. |
| [ExcelApi 1.7](excel-api-1-7-requirement-set.md)  | Версия 1801 (сборка 9001.2171) или более поздняя   | 2.9 или более поздняя  | 16.9 или более поздняя  | Апрель 2018 г. |
| [ExcelApi 1.6](excel-api-1-6-requirement-set.md)  | Версия 1704 (сборка 8201.2001) или более поздняя   | Версия 2.2 или более поздняя  | Версия 15.36 или более поздняя | Апрель 2017 г. |
| [ExcelApi 1.5](excel-api-1-5-requirement-set.md)  | Версия 1703 (сборка 8067.2070) или более поздняя   | Версия 2.2 или более поздняя  | Версия 15.36 или более поздняя | Март 2017 г. |
| [ExcelApi 1.4](excel-api-1-4-requirement-set.md)  | Версия 1701 (сборка 7870.2024) или более поздняя   | Версия 2.2 или более поздняя  | Версия 15.36 или более поздняя | Январь 2017 г. |
| [ExcelApi 1.3](excel-api-1-3-requirement-set.md)  | Версия 1608 (сборка 7369.2055) или выше   | 1.27 или более поздняя | 15.27 или более поздняя | Сентябрь 2016 г. |
| [ExcelApi 1.2](excel-api-1-2-requirement-set.md)  | Версия 1601 (сборка 6741.2088) или выше   | 1.21 или более поздняя | 15.22 или более поздняя | Январь 2016 г. |
| [ExcelApi 1.1](excel-api-1-1-requirement-set.md)  | Версия 1509 (сборка 4266.1001) или более поздняя   | 1.19 или более поздняя | 15.20 или более поздняя | Январь 2016 г. |

> [!NOTE]
> Бессрочные версии Office поддерживают следующие наборы обязательных элементов:
>
> - Office 2019 поддерживает ExcelApi 1.8 и более ранние версии.
> - Office 2016 поддерживает только набор обязательных элементов ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание ведущих приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)
