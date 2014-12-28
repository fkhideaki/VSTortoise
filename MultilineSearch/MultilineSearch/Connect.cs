using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Resources;
using System.Reflection;
using System.Globalization;

namespace MultilineSearch
{
	/// <summary>アドインを実装するためのオブジェクトです。</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>アドイン オブジェクトのコンストラクターを実装します。初期化コードをこのメソッド内に配置してください。</summary>
		public Connect()
		{
		}

		/// <summary>IDTExtensibility2 インターフェイスの OnConnection メソッドを実装します。アドインが読み込まれる際に通知を受けます。</summary>
		/// <param term='application'>ホスト アプリケーションのルート オブジェクトです。</param>
		/// <param term='connectMode'>アドインの読み込み状態を説明します。</param>
		/// <param term='addInInst'>このアドインを表すオブジェクトです。</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
			if(connectMode == ext_ConnectMode.ext_cm_UISetup)
			{
				object []contextGUIDS = new object[] { };
				Commands2 commands = (Commands2)_applicationObject.Commands;

				//コマンドを [ツール] メニューに配置します。
				//メイン メニュー項目のすべてを保持するトップレベル コマンド バーである、MenuBar コマンド バーを検索します:
				Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar;
				menuBarCommandBar = ((Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars)["MenuBar"];

				//MenuBar コマンド バーで [ツール] コマンド バーを検索します:
				CommandBarControls controls = menuBarCommandBar.Controls;
				CommandBarControl toolsControl = controls["Tools"];
				CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

				//アドインによって処理する複数のコマンドを追加する場合、この try ブロックおよび catch ブロックを重複できます。
				//  ただし、新しいコマンド名を含めるために QueryStatus メソッドおよび Exec メソッドの更新も実行してください。
				try
				{
					//コマンド コレクションにコマンドを追加します:
					Command command = commands.AddNamedCommand2(
						_addInInstance,
						"MultilineSearch",
						"MultilineSearch",
						"Executes the command for MultilineSearch",
						true,
						59,
						ref contextGUIDS,
						(int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled,
						(int)vsCommandStyle.vsCommandStyleText,
						vsCommandControlType.vsCommandControlTypeButton);

					//コマンドのコントロールを [ツール] メニューに追加します:
					if((command != null) && (toolsPopup != null))
					{
						command.AddControl(toolsPopup.CommandBar, 1);
					}
				}
				catch(System.ArgumentException)
				{
					//同じ名前のコマンドが既に存在しているため、例外が発生した可能性があります。
					//  その場合、コマンドを再作成する必要はありません。 例外を 
					//  無視しても安全です。
				}
			}
		}

		/// <summary>IDTExtensibility2 インターフェイスの OnDisconnection メソッドを実装します。アドインがアンロードされる際に通知を受けます。</summary>
		/// <param term='disconnectMode'>アドインのアンロード状態を説明します。</param>
		/// <param term='custom'>ホスト アプリケーション固有のパラメーターの配列です。</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 インターフェイスの OnAddInsUpdate メソッドを実装します。アドインのコレクションが変更されたときに通知を受けます。</summary>
		/// <param term='custom'>ホスト アプリケーション固有のパラメーターの配列です。</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 インターフェイスの OnStartupComplete メソッドを実装します。ホスト アプリケーションが読み込みを終了したときに通知を受けます。</summary>
		/// <param term='custom'>ホスト アプリケーション固有のパラメーターの配列です。</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>IDTExtensibility2 インターフェイスの OnBeginShutdown メソッドを実装します。ホスト アプリケーションがアンロードされる際に通知を受けます。</summary>
		/// <param term='custom'>ホスト アプリケーション固有のパラメーターの配列です。</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}
		
		/// <summary>IDTCommandTarget インターフェイスの QueryStatus メソッドを実装します。これは、コマンドの可用性が更新されたときに呼び出されます。</summary>
		/// <param term='commandName'>状態を決定するためのコマンド名です。</param>
		/// <param term='neededText'>コマンドに必要なテキストです。</param>
		/// <param term='status'>ユーザー インターフェイス内のコマンドの状態です。</param>
		/// <param term='commandText'>neededText パラメーターから要求されたテキストです。</param>
		/// <seealso class='Exec' />
		public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus status, ref object commandText)
		{
			if(neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
				if(commandName == "MultilineSearch.Connect.MultilineSearch")
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported|vsCommandStatus.vsCommandStatusEnabled;
					return;
				}
			}
		}

		/// <summary>IDTCommandTarget インターフェイスの Exec メソッドを実装します。これは、コマンドが実行されるときに呼び出されます。</summary>
		/// <param term='commandName'>実行するコマンド名です。</param>
		/// <param term='executeOption'>コマンドの実行方法を説明します。</param>
		/// <param term='varIn'>呼び出し元からコマンド ハンドラーへ渡されたパラメーターです。</param>
		/// <param term='varOut'>コマンド ハンドラーから呼び出し元へ渡されたパラメーターです。</param>
		/// <param term='handled'>コマンドが処理されたかどうかを呼び出し元に通知します。</param>
		/// <seealso class='Exec' />
		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if(executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
				if(commandName == "MultilineSearch.Connect.MultilineSearch")
				{
					FindMultiline f = new FindMultiline(_applicationObject);
					f.Execute();
					handled = true;
					return;
				}
			}
		}
		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}