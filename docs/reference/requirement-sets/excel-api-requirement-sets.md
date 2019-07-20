---
title: Наборы обязательных элементов API JavaScript для Excel
description: Сведения о наборе обязательных элементов надстройки Office для сборок Excel
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: ec49cbdadf65b653170f9b5cbcafa6aaf0fa5177
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804978"
---
# <a name="excel-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Excel

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Доступность набора обязательных элементов

Надстройки Excel работают в нескольких версиях Office, включая Office 2016 или более поздние версии для Windows, а также Office в Интернете, Office для Mac и Office для iPad. В приведенной ниже таблице перечислены наборы обязательных элементов для Excel, ведущие приложения Office, которые поддерживают все наборы обязательных инструментов, а также номера сборок или версий для этих приложений.

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов, следует ссылаться на **рабочую** библиотеку в сети CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Сведения об использовании API предварительных версий см. в статье [Предварительные версии API JavaScript для Excel](./excel-preview-apis.md).

|  Набор обязательных элементов  |  Office для Windows<br>(версия, подключенная к подписке на Office 365)  |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете |
|:-----|-----|:-----|:-----|:-----|:-----|
| [Предварительная версия](excel-preview-apis.md)  | Применяйте последнюю версию Office для использования предварительных версий API (может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider)) |
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
> Номер сборки Office 2016, установленной с помощью MSI, — 16.0.4266.1001. Эта версия содержит только набор обязательных элементов ExcelApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
- [Обзор Office Online Server](/officeonlineserver/office-online-server-overview)
