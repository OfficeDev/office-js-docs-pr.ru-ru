---
title: Тестирование надстроек Office
description: Узнайте, как протестировать надстройку Office.
ms.date: 07/28/2022
ms.localizationpriority: high
ms.openlocfilehash: 56052182eafae59d42044ce4be40e086e51e8103
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467246"
---
# <a name="test-office-add-ins"></a>Тестирование надстроек Office

Эта статья содержит рекомендации по тестированию, отладке и диагностике надстроек Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Тестирование кроссплатформенных выпусков и выпусков для нескольких версий Office

Надстройки Office запускаются на основных платформах, поэтому требуется протестировать надстройку на всех платформах, где ваши пользователи могут запускать Office. Обычно это относится Office в Интернете, Office для Windows (бессрочная и подписка На Microsoft 365), Office для Mac, Office на iOS и (для надстроек Outlook) Office на Android. Однако могут возникать ситуации, когда вы точно знаете, что никто из ваших пользователей не будет работать на некоторых платформах. Например, если вы создаете надстройку для компании, которой требуется, чтобы ее пользователи могли работать с компьютерами Windows и подпиской на Office, вам не нужно тестировать Office на Mac или бессрочное использование Office в Windows.

> [!NOTE]
> На компьютерах с Windows браузер, используемый надстройкой, определяется версией Windows и Office. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Надстройки, предлагаемые через AppSource, проходят проверку, включающую тестирование на всех платформах. Кроме того, надстройки тестируются в Office для Интернета со всеми основными современными браузерами, включая Microsoft Edge (WebView2 на основе Chromium), Chrome и Safari. Соответственно, перед отправкой в AppSource необходимо протестировать эти платформы и браузеры. Дополнительные сведения о проверке см. в статье [Политики сертификации коммерческой платформы Marketplace](/legal/marketplace/certification-policies), особенно в [разделе 1120.3](/legal/marketplace/certification-policies#11203-functionality), а также на странице [Доступность и применение надстроек Office](/javascript/api/requirement-sets).
>
> AppSource не использует Internet Explorer или устаревшую версию Microsoft Edge (WebView1) для тестирования надстроек в Office для Интернета. Но если значительное число ваших пользователей будет использовать браузер Edge прежних версий для открытия Office в Интернете, вам следует протестировать надстройку с ним. (Office в Интернете не будет открываться в Internet Explorer, поэтому тестировать надстройку с этим браузером не нужно.) Дополнительные сведения см. в статьях "[Поддержка Internet Explorer 11](../develop/support-ie-11.md)" и "[Устранение неполадок Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshoot-microsoft-edge-issues)". Office по-прежнему поддерживает эти браузеры как поставщики сред выполнения надстроек, поэтому если вы считаете, что столкнулись с ошибкой в работе надстроек в них, создайте запись о проблеме для репозитория [office-js.](https://github.com/OfficeDev/office-js/issues/new/choose)

## <a name="sideload-an-office-add-in-for-testing"></a>Загрузка неопубликованной надстройки Office для тестирования

You can use sideloading to install an Office Add-in for testing without having to first put it in an add-in catalog. The procedure for sideloading an add-in varies by platform, and in some cases, by product as well. The following articles each describe how to sideload Office Add-ins on a specific platform or within a specific product.

- [Загрузка неопубликованных надстроек Office в Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Загрузка неопубликованных надстроек Office в Office в Интернете](sideload-office-add-ins-for-testing.md)

- [Загрузка неопубликованных надстроек Office на Mac](sideload-an-office-add-in-on-mac.md)

- [Загрузка неопубликованных надстроек Office на iPad](sideload-an-office-add-in-on-ipad.md)

- [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="unit-testing"></a>Модульное тестирование

Сведения о том, как добавить модульные тесты в проект надстройки, см. в статье [Модульное тестирование в надстройках Office](unit-testing.md).

## <a name="debug-an-office-add-in"></a>Отладка надстройки Office

Процедура отладки надстройки Office зависит от вашей платформы и среды. Дополнительные сведения см. в статье [Отладка надстроек Office](debug-add-ins-overview.md).

## <a name="validate-an-office-add-in-manifest"></a>Проверка манифеста надстройки Office

Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Устранение ошибок, с которыми сталкиваются пользователи

Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).
