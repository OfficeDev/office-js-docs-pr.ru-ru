---
title: Наборы обязательных элементов API JavaScript для Word
description: Сведения о наборе обязательных элементов надстройки Office для сборок Word.
ms.date: 10/08/2021
ms.prod: word
ms.localizationpriority: high
ms.openlocfilehash: 88d49b860cb832f59e8e2a631f0c4a3c57698306
ms.sourcegitcommit: a37be80cf47a37c85b7f5cab216c160f4e905474
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/09/2021
ms.locfileid: "60250534"
---
# <a name="word-javascript-api-requirement-sets"></a>Наборы обязательных элементов API JavaScript для Word

Наборы обязательных элементов — именованные группы элементов API. Надстройки Office с помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения определяют, поддерживает ли приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md).

## <a name="requirement-set-availability"></a>Доступность набора обязательных элементов

Надстройки Word работают в нескольких версиях Office, включая Office 2016 для Windows, Office в Интернете, для iPad и Mac. В приведенной ниже таблице перечислены наборы требований Word, ведущие приложения Office, которые их поддерживают, а также номера сборок или версий для этих приложений.

> [!NOTE]
> Чтобы использовать API в любом из нумерованных наборов обязательных элементов, следует ссылаться на **рабочую** библиотеку в сети CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.
>
> Сведения об использовании API предварительных версий см. в статье [Предварительные версии API JavaScript для Word](word-preview-apis.md).

|  Набор обязательных элементов  |   Office для Windows\*<br>(подключено к подписке на Microsoft 365)  |  Office для iPad<br>(подключено к подписке на Microsoft 365)  |  Office для Mac<br>(подключено к подписке на Microsoft 365)  | Office в Интернете  |
|:-----|-----|:-----|:-----|:-----|
| [Предварительная версия](word-preview-apis.md) | Применяйте последнюю версию Office для использования предварительных версий API (может потребоваться присоединение к [программе предварительной оценки Office](https://insider.office.com)) |
| [WordApi 1.3](word-api-1-3-requirement-set.md) | Версия 1612 (сборка 7668.1000) или более поздняя| Март 2017 г., 2.22 или более поздняя | Март 2017 г., 15.32 или более поздняя| Март 2017 г. |
| [WordApi 1.2](word-api-1-2-requirement-set.md) | Обновление за декабрь 2015 г., версия 1601 (сборка 6568.1000) или выше | Январь 2016 г., версия 1.18 или выше | Январь 2016 г., версия 15.19 или выше| Сентябрь 2016 г. |
| [WordApi 1.1](word-api-1-1-requirement-set.md) | Версия 1509 (сборка 4266.1001) или более поздняя| Январь 2016 г., версия 1.18 или выше | Январь 2016 г., версия 15.19 или выше| Сентябрь 2016 г. |

> [!NOTE]
> Бессрочные версии Office поддерживают следующие наборы обязательных элементов:
>
> - Office 2019 и Office 2021 поддерживают WordApi 1.3 и более ранние версии.
> - Office 2016 поддерживает только набор обязательных элементов WordApi 1.1.

## <a name="office-versions-and-build-numbers"></a>Номера версий и сборок Office

Дополнительные сведения о номерах версий и сборок Office см. в следующих статьях:

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Версии Office и наборы обязательных элементов](../../develop/office-versions-and-requirement-sets.md)
- [Указание приложений Office и обязательных элементов API](../../develop/specify-office-hosts-and-api-requirements.md)
- [XML-манифест надстроек Office](../../develop/add-in-manifests.md)
