import tkinter
import pytest


def test_真实窗口读取控件配置且没有重复启动(tmp_path,monkeypatch):
    from src.界面 import App
    try:
        app=App(root_dir=tmp_path)
    except tkinter.TclError:
        pytest.skip('当前环境没有可用Windows桌面会话。')
    app.withdraw()
    try:
        app.url_entry.delete(0,'end');app.url_entry.insert(0,'http://localhost:8000/v1')
        app.key_entry.insert(0,'测试密钥')
        app.model_entry.delete(0,'end');app.model_entry.insert(0,'窗口模型')
        cfg=app.read_form()
        assert cfg['base_url']=='http://localhost:8000/v1'
        assert cfg['model']=='窗口模型' and cfg['api_key']=='测试密钥'
        assert app.key_entry.cget('show')!=''
        monkeypatch.setattr('src.界面.messagebox.showwarning',lambda *a,**k:None)
        app.start_button.invoke()
        assert not app.busy
        app.update_idletasks()
    finally:
        app.destroy()