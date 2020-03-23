@ECHO OFF
echo 复制数据
if exist %~dp0..\Assets\AssetResources\Config ( echo "Config文件存在") else ( md %~dp0..\Assets\AssetResources\Config )
for %%i in (%~dp0..\Assets\AssetResources\Config\*.bytes) do ( 
    del %%i
)
for %%i in (%~dp0Example\Export\dat\*.dat) do ( 
    copy /y %%i %~dp0..\Assets\AssetResources\Config
)
rename %~dp0..\Assets\AssetResources\Config\*.dat *.dat.bytes
echo 复制数据完毕
echo 复制cs文件
for %%i in (%~dp0Example\Export\language\csharpForILRumtime\*.cs) do ( 
    copy /y %%i %~dp0..\Assets\Scripts\UnityHotfix\Protobuf
)
echo 复制cs文件完毕
pause