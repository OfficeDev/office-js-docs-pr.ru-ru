---
title: Установка последней версии Office
description: Сведения о том, как получать последние сборки Office раньше других.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: d72012e3c2e642c74d8573c4d9bb3b29a8fc0274
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076022"
---
# <a name="install-the-latest-version-of-office"></a>Установка последней версии Office

Первыми новые функции для разработчиков, в том числе предварительные версии, получают подписчики, которые получают последние сборки Office раньше других.

## <a name="opt-in-to-getting-the-latest-builds"></a>Как получать последние сборки раньше других

Чтобы получать последние сборки Office раньше других:

- Если вы абонент Microsoft 365 для семьи, личный или университетский, см. статью [Be an Office Insider](https://insider.office.com).
- Если вы клиент Приложения Microsoft 365 для бизнеса, см. в выпуске [Install the First Release build for Приложения Microsoft 365 для бизнеса клиентов.](https://support.office.com/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead)
- Если вы используете Office для Mac:
  - Запустите приложение Office.
  - Выберите пункт **Проверить наличие обновлений** в меню "Справка".
  - В окне "Автоматическое обновление (Майкрософт)" установите флажок для участия в программе предварительной оценки Office.

## <a name="get-the-latest-build"></a>Как получить последнюю сборку

Чтобы получить последнюю сборку Office:

1. Скачайте [средство развертывания Office](https://www.microsoft.com/download/details.aspx?id=49117).
2. Запустите это средство. Будут извлечены два файла: Setup.exe и configuration.xml.
3. Замените файл configuration.xml [файлом конфигурации первого выпуска](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml).
4. Выполните следующую команду от имени администратора: `setup.exe /configure configuration.xml`

> [!NOTE]
> Команда может выполняться долго, при этом ход ее выполнения нигде не отображается.

По завершении процесса установки у вас будут последние версии приложений Office. Чтобы убедиться, что у вас последняя сборка, в любом приложении Office последовательно выберите **Файл** > **Учетная запись**. В разделе "Обновления Office" над номером версии должна быть надпись "Предварительная оценка Office".

![Снимок экрана, на который показаны сведения о продукте с Office insiders.](../images/office-insiders-label.png)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Минимальные сборки Office, которые могут использовать наборы обязательных элементов API JavaScript для Office

Сведения о минимальных сборках продуктов для каждой платформы см. в следующих статьях:

- [Наборы обязательных элементов API JavaScript для Excel](../reference/requirement-sets/excel-api-requirement-sets.md)
- [Наборы обязательных элементов API JavaScript для OneNote](../reference/requirement-sets/onenote-api-requirement-sets.md)
- [Наборы обязательных элементов API JavaScript для Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md)
- [Наборы обязательных элементов API JavaScript для PowerPoint](../reference/requirement-sets/powerpoint-api-requirement-sets.md)
- [Наборы обязательных элементов API JavaScript для Word](../reference/requirement-sets/word-api-requirement-sets.md)
- [Наборы обязательных элементов API диалоговых окон](../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Наборы обязательных элементов общего API для Office](../reference/requirement-sets/office-add-in-requirement-sets.md)
