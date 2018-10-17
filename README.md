# USTC-kms

**用于激活Windows和Office组件的kms脚本**

基于[官方激活脚本](http://zbh.ustc.edu.cn/msiso/UstcKms.bat)进行修改，对Office的激活过程进行了以下优化

* 不再需要右键管理员运行，现在双击运行脚本可以自动申请管理员权限啦(๑•̀ㅂ•́)و✧

* 原来的脚本要求Office安装在默认位置，否则用户必须手动修改脚本内容，现在新脚本支持自动从注册表查找Office安装位置啦(๑•̀ㅂ•́)و✧



**注意：**

1. **和官方脚本一样，本脚本也仅能激活来自[中科大正版软件](http://zbh.ustc.edu.cn/zbh.php)的发行版，且仍然要处于校园网环境内。**
2. **注册表查找基于Word的路径，所以已安装套件中必须包含Word，如果安装Office时未安装Word，本脚本将会无法工作。**



>已通过测试的Office版本包括：
>
>Office-2016-x64
>
>Office-2013-x32
>
>Office-2010-x64
