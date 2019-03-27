---
title: Загрузка неопубликованных надстроек Office с помощью специальной команды
description: ''
ms.date: 03/19/201907/24/2018
localization_priority: Priority
ms.openlocfilehash: dfa231374133ad857554afaf343362f1415788f4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870116"
---
# <a name="sideload-office-add-ins-for-testing-using-the-sideload-command"></a>Загрузка неопубликованных надстроек Office для тестирования с помощью **специальной команды**
 >[!NOTE]
>Метод "npm run sideload" работает только для надстроек Excel, Word и PowerPoint, которые запускаются в Windows, и только для проектов надстройки, созданных с помощью [инструмента **yo office**](https://github.com/OfficeDev/generator-office), и у которых есть скрипт `sideload` в разделе `scripts` файла package.json. (У проектов, созданных с помощью более ранних версий **yo office** также нет этого скрипта). Если ваш проект был создан с помощью Visual Studio или у него нет скрипта загрузки неопубликованных приложений, вы можете загрузить его неопубликованным в Windows с помощью метода, описанного в статье [Загрузка неопубликованной надстройки Office из общей сетевой папки](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
>
> Если вы не тестируете надстройку Word, Excel или PowerPoint в Windows, см. одну из следующих статей:
> 
> - [Загрузка неопубликованных надстроек Office в Office Online для тестирования](sideload-office-add-ins-for-testing.md)
> - [Загрузка неопубликованных надстроек Office на iPad и Mac для тестирования](sideload-an-office-add-in-on-ipad-and-mac.md)
> - [Загрузка неопубликованных надстроек Outlook для тестирования](/outlook/add-ins/sideload-outlook-add-ins-for-testing)

1. Откройте командную строку от имени администратора.

2. Измените каталоги на корневой каталог папки вашего проекта надстройки.

3. Выполните следующую команду, чтобы запустить экземпляр локального сервера на порту 3000 для обслуживания вашего проекта надстройки: "**npm run start**"

4. Откройте вторую командную строку от имени администратора.

5. Измените каталоги на корневой каталог папки вашего проекта надстройки.

6. Выполните следующую команду для загрузки ведущего приложения (например, Excel или Word) и регистрации надстройки в ведущем приложении: "**npm run sideload**"

## <a name="see-also"></a>См. также

- [Проверка манифеста и устранение связанных с ним неполадок](troubleshoot-manifest.md)
- [Публикация надстройки Office](../publish/publish.md)
