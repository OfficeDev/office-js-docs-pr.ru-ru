---
title: Элемент MobileFormFactor в файле манифеста
description: Элемент MobileFormFactor указывает параметры параметров формы мобильного устройства для надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 64a7681ca23becf42af1ba435aae4d509e6ad1ba
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612229"
---
# <a name="mobileformfactor-element"></a>Элемент MobileFormFactor

Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.

Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** . Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).

Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.

## <a name="child-elements"></a>Дочерние элементы

| Элемент                               | Обязательный | Описание  |
|:--------------------------------------|:--------:|:-------------|
| [ExtensionPoint](extensionpoint.md) | Да      | Определяет, где предоставляются функции надстройки. |
| [FunctionFile](functionfile.md)     | Да      | URL-адрес файла, который содержит функции JavaScript.|

## <a name="mobileformfactor-example"></a>Пример MobileFormFactor

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
