全角・半角の校閲を行います。
コマンドライン上で操作を行なってください。

# 以下を入力してプログラムが存在するディレクトリに移動してください。
cd /home/kita/zenkaku_hankaku_check/

# venv仮想環境を作成してください。
python3 -m venv venv
※venvは任意の名前を設定できます。その場合は python3 -m venv projectenv のように変更してください。

# venv仮想環境を有効化してください。
source venv/bin/activate
※venvに任意の名前を設定した場合は source projectenv/bin/activate のように変更してください。

# dataディレクトリに校閲対象のファイルを格納してください。

# 以下を入力してプログラムを実行してください。
python main.py

# (オプション)以下を入力するとプログラム実行時に生成したファイルを一括で削除できます。
python delete_files.py
