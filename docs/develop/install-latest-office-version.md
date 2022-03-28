---
title: Установка последней версии Office
description: Сведения о том, как получать последние сборки Office раньше других.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: c558da4540638c91ed3519685de007379d1e1061
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483663"
---
# <a name="install-the-latest-version-of-office"></a>Установка последней версии Office

Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других.

## <a name="opt-in-to-getting-the-latest-builds-of-office"></a>Выбор в получении последних сборки Office

- Если вы абонент Microsoft 365 для семьи, personal или university, см. статью [Be an Office Insider](https://insider.office.com).
- Если вы клиент Приложения Microsoft 365 для бизнеса, см. в этой версии Сборка первого выпуска для [Приложения Microsoft 365 для бизнеса клиентов](https://support.office.com/article/4dd8ba40-73c0-4468-b778-c7b744d03ead).
- Если вы используете Office для Mac:
  - Запустите приложение Office.
  - Выберите пункт **Проверить наличие обновлений** в меню "Справка".
  - В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office.

## <a name="get-the-latest-build-of-office"></a>Получите последнюю сборку Office

1. Скачайте [средство развертывания Office](https://www.microsoft.com/download/details.aspx?id=49117).
2. Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.
3. Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml`

> [!NOTE]
> Команда может выполняться долго, при этом ход ее выполнения нигде не отображается.

По завершении процесса установки у вас будут последние версии приложений Office. Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл** > **Учетная запись**. В разделе "Обновления Office" над номером версии должна быть надпись "Предварительная оценка Office".

![Снимок экрана, на который показаны сведения о продукте с Office insiders.](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office

- [Наборы обязательных элементов API JavaScript для Excel](/javascript/api/requirement-sets/excel-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для OneNote](/javascript/api/requirement-sets/onenote-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для Outlook](/javascript/api/requirement-sets/outlook-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для PowerPoint](/javascript/api/requirement-sets/powerpoint-api-requirement-sets)
- [Наборы обязательных элементов API JavaScript для Word](/javascript/api/requirement-sets/word-api-requirement-sets)
- [Наборы обязательных элементов API диалоговых окон](/javascript/api/requirement-sets/dialog-api-requirement-sets)
- [Наборы обязательных элементов общего API для Office](/javascript/api/requirement-sets/office-add-in-requirement-sets)
