---
title: Тестирование надстроек Office
description: Узнайте, как протестировать надстройку Office
ms.date: 12/02/2021
ms.localizationpriority: high
ms.openlocfilehash: 8d57f396c5387faf22ba8b03fd2e5019be4e14d2
ms.sourcegitcommit: 33824aa3995a2e0bcc6d8e67ada46f296c224642
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/12/2022
ms.locfileid: "61765915"
---
# <a name="test-office-add-ins"></a>Тестирование надстроек Office

Эта статья содержит рекомендации по тестированию, отладке и диагностике надстроек Office.

## <a name="test-cross-platform-and-for-multiple-versions-of-office"></a>Тестирование кроссплатформенных выпусков и выпусков для нескольких версий Office

Надстройки Office запускаются на основных платформах, поэтому требуется протестировать надстройку на всех платформах, где ваши пользователи могут запускать Office. К ним обычно относятся Office в Интернете, Office для Windows (как подписка, так и единовременная покупка), Office для Mac, Office для iOS и (для надстроек Outlook) Office для Android. Однако могут возникать ситуации, когда вы точно знаете, что никто из ваших пользователей не будет работать на некоторых платформах. Например, если вы создаете надстройку для компании, которая требует, чтобы пользователи работали на компьютерах с Windows и подпиской на Office, вам не нужно выполнять тестирование в Office для Mac или единовременно приобретенных экземплярах для Windows.

> [!NOTE]
> На компьютерах с Windows браузер, используемый надстройкой, определяется версией Windows и Office. Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).

> [!IMPORTANT]
> Надстройки, предлагаемые через AppSource, проходят проверку, включающую тестирование на всех платформах. Кроме того, надстройки тестируются в Office для Интернета со всеми основными современными браузерами, включая Microsoft Edge (WebView2 на основе Chromium), Chrome и Safari. Соответственно, перед отправкой в AppSource необходимо протестировать эти платформы и браузеры. Дополнительные сведения о проверке см. в статье [Политики сертификации коммерческой платформы Marketplace](/legal/marketplace/certification-policies), особенно в [разделе 1120.3](/legal/marketplace/certification-policies#11203-functionality), а также на странице [Доступность и применение надстроек Office](../overview/office-add-in-availability.md).
>
> AppSource не использует Internet Explorer или устаревшую версию Microsoft Edge (WebView1) для тестирования надстроек в Office для Интернета. Но если значительное число ваших пользователей будет использовать браузер Edge прежних версий для открытия Office в Интернете, вам следует протестировать надстройку с ним. (Office в Интернете не будет открываться в Internet Explorer, поэтому тестировать надстройку с этим браузером не нужно.) Дополнительные сведения см. в статьях "[Поддержка Internet Explorer 11](../develop/support-ie-11.md)" и "[Устранение неполадок Microsoft Edge](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues)". Office по-прежнему поддерживает эти браузеры как поставщики сред выполнения надстроек, поэтому если вы считаете, что столкнулись с ошибкой в работе надстроек в них, создайте запись о проблеме для репозитория [office-js.](https://github.com/OfficeDev/office-js/issues/new/choose)

## <a name="sideload-an-office-add-in-for-testing"></a>Загрузка неопубликованной надстройки Office для тестирования

Вы можете установить надстройку Office для тестирования, не размещая ее в каталоге надстроек. Процедура отличается для разных платформ, а в некоторых случаях и для разных продуктов. Следующие статьи посвящены загрузке неопубликованных надстроек Office на определенной платформе или в определенном продукте.

- [Загрузка неопубликованных надстроек Office в Windows](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

- [Загрузка неопубликованных надстроек Office в Office в Интернете](sideload-office-add-ins-for-testing.md)

- [Загрузка неопубликованных надстроек Office на iPad и Mac](sideload-an-office-add-in-on-ipad-and-mac.md)

- [Загрузка неопубликованных надстроек Outlook для тестирования](../outlook/sideload-outlook-add-ins-for-testing.md)

## <a name="unit-testing"></a>Модульное тестирование

Сведения о том, как добавить модульные тесты в проект надстройки, см. в статье [Модульное тестирование в надстройках Office](unit-testing.md).

## <a name="debug-an-office-add-in"></a>Отладка надстройки Office

Процедура отладки надстройки Office зависит от вашей платформы и среды. Дополнительные сведения см. в статье [Отладка надстроек Office](debug-add-ins-overview.md).

## <a name="validate-an-office-add-in-manifest"></a>Проверка манифеста надстройки Office

Информацию о проверке манифеста надстройки Office и устранении связанных с ним неполадок см. в [этой статье](troubleshoot-manifest.md).

## <a name="troubleshoot-user-errors"></a>Устранение ошибок, с которыми сталкиваются пользователи

Информацию об устранении основных ошибок, с которыми сталкиваются пользователи при работе с надстройками Office, см. в [этой статье](testing-and-troubleshooting.md).
