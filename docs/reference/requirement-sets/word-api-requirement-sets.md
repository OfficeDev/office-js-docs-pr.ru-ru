---
title: Наборы обязательных элементов API JavaScript для Word
description: Сведения о наборе обязательных элементов надстройки Office для сборок Word
ms.date: 01/06/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: c90daafe46d301b404ee902b38bb7417562adc44
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969533"
---
# <a name="word-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Word

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets).

## <a name="requirement-set-availability"></a>Доступность набора обязательных элементов

Надстройки Word работают в нескольких версиях Office, включая Office 2016 или более поздней версии для Windows, а также Office в Интернете, Office для iPad и Office для Mac. В приведенной ниже таблице перечислены наборы требований Word, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов, следует ссылаться на **рабочую** библиотеку в сети CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Сведения об использовании API предварительных версий см. в статье [Предварительные версии API JavaScript для Excel](word-preview-apis.md).

|  Набор обязательных элементов  |   Office для Windows\*<br>(версия, подключенная к подписке на Office 365)  |  Office для iPad<br>(версия, подключенная к подписке на Office 365)  |  Office для Mac<br>(версия, подключенная к подписке на Office 365)  | Office в Интернете  |
|:-----|-----|:-----|:-----|:-----|
| [Предварительная версия](word-preview-apis.md) | Применяйте последнюю версию Office для использования предварительных версий API (может потребоваться присоединение к [программе предварительной оценки Office](https://products.office.com/office-insider)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Версия 1612 (сборка 7668.1000) или более поздняя| Март 2017 г., 2.22 или более поздняя | Март 2017 г., 15.32 или более поздняя| Март 2017 г. |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Обновление за декабрь 2015 г., версия 1601 (сборка 6568.1000) или выше | Январь 2016 г., версия 1.18 или выше | Январь 2016 г., версия 15.19 или выше| Сентябрь 2016 г. |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Версия 1509 (сборка 4266.1001) или более поздняя| Январь 2016 г., версия 1.18 или выше | Январь 2016 г., версия 15.19 или выше| Сентябрь 2016 г. |

> [!NOTE]
> Бессрочные версии Office поддерживают следующие наборы обязательных элементов:
>
> - Office 2019 поддерживает WordApi 1.3 и более ранние версии.
> - Office 2016 поддерживает только набор обязательных элементов WordApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

- [Номера версий и сборок выпусков из канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);
- [Какая у меня версия Office](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19);
- [Где можно найти номера версии и сборки клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7);

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Версии Office и наборы обязательных элементов](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](/office/dev/add-ins/develop/add-in-manifests)
