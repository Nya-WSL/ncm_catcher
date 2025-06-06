# ncm_catcher
从Windows窗口标题获取PC端网易云音乐当前播放歌名并支持写入文件，可通过OBS读取

# Build

```shell
pip install poetry
poetry install
poetry run pyinstaller -F main.py --name ncm_catcher
```
