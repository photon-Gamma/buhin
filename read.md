# 環境構築
インストール
```
sudo chown koki ~/.bash_profile
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
```

設定の確認
```
rustup update
rustup show
rustup default stable
```

コンパイルと実行
```
rustc ./src/main.rs
./main
```
# プログラミング

## パッケージの作成
cargoを使うことでRustパッケージを作成できます。パッケージとは、ソースコードやマニフェストファイル（Cargo.toml）をまとめたものです。パッケージには1つ以上のクレートが含まれています。クレートはコードをまとめたもので、ビルドするとバイナリやライブラリを作成します。パッケージは、バイナリクレートまたはライブラリクレート、サンプル、テスト、ベンチマークなどのターゲットを持っています。ターゲットはそれぞれ何かしらの成果物を出力します。Rustにおいて、プロジェクトはパッケージの別名と考えることができます。パッケージを作成する方法は2種類あり、cargo initコマンドとcargo newコマンドです。まずはcargo initコマンドについて説明します。これはカレントディレクトリ内に最低限のパッケージ構成を作成します。
```
cargo init
```

標準でバイナリクレートをターゲットにした構成が作成されます。この状態でcargo buildコマンドを実行するとバイナリクレートをビルドします。
```
cargo build
```

cargo runコマンドで生成されたバイナリを実行します。ビルドせずにcargo runコマンドを実行すると自動的にビルドされます。
```
cargo run
```

## パッケージの作成
最小構成の構成を作成する
```
cargo init test2
```
test2のディレクトリに最小限のパッケージが作成される。ここに、sourceフォルダを追加する。
```
cargo build hello
cargo run --source hello
```