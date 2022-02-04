---
title: Элемент MobileFormFactor в файле манифеста
description: Элемент MobileFormFactor указывает параметры мобильного форм-фактора для надстройки.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="mobileformfactor-element"></a>Элемент MobileFormFactor

Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Каждое определение MobileFormFactor** содержит элемент **FunctionFile** и один или несколько **элементов ExtensionPoint**. Дополнительные сведения см. в [элементе FunctionFile и](functionfile.md) [элементе ExtensionPoint](extensionpoint.md).

Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.

## <a name="child-elements"></a>Дочерние элементы

| Элемент                             | Обязательный | Описание  |
|:------------------------------------|:--------:|:-------------|
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
