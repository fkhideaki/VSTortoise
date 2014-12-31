using EnvDTE;
using EnvDTE80;
using Extensibility;
using TortoiseCore;
using Microsoft.VisualStudio.CommandBars;
using System;


namespace TortoiseAddin
{
	/// <summary>アドインを実装するためのオブジェクトです。</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2, IDTCommandTarget
	{
		/// <summary>アドイン オブジェクトのコンストラクターを実装します。初期化コードをこのメソッド内に配置してください。</summary>
		public Connect()
		{
		}

		bool IsInitializeEvent(ext_ConnectMode connectMode)
		{
			if (connectMode == ext_ConnectMode.ext_cm_UISetup)
				return true;

			//if (connectMode == ext_ConnectMode.ext_cm_AfterStartup)
			//	return true;

			//if (connectMode == ext_ConnectMode.ext_cm_Startup)
			//	return true;

			return false;
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
			if (!IsInitializeEvent(connectMode))
				return;

			object[] contextGUIDS = new object[] { };
			Commands2 commands = (Commands2)_applicationObject.Commands;

			//コマンドを [ツール] メニューに配置します。
			//メイン メニュー項目のすべてを保持するトップレベル コマンド バーである、MenuBar コマンド バーを検索します:
			Microsoft.VisualStudio.CommandBars.CommandBars command_bars;
			command_bars = (Microsoft.VisualStudio.CommandBars.CommandBars)_applicationObject.CommandBars;
			Microsoft.VisualStudio.CommandBars.CommandBar menuBarCommandBar = (command_bars)["MenuBar"];

			// アドインインストール直後しか正しくメニューをつくれないっぽい.
			// たぶん
			// 1.初回起動時なら→コマンドを作成.
			// 2.コマンドをメニューに登録する.
			// という流れが正しいがやり方がわからない.
			string menu_root = "TortoiseAddin (&A)";
			CommandBarControl toolsControl = null;
			try
			{
				toolsControl = menuBarCommandBar.Controls[menu_root];
			}
			catch (Exception)
			{
				toolsControl = menuBarCommandBar.Controls.Add(Type: Microsoft.VisualStudio.CommandBars.MsoControlType.msoControlPopup);
				toolsControl.Caption = menu_root;
			}

			CommandBarPopup toolsPopup = (CommandBarPopup)toolsControl;

			CommandList cmd_list = CreateOrGetCommandDef();
			for (int i = 0; i < cmd_list.NumCommands(); i++)
			{
				CreateCommand(ref contextGUIDS, commands, toolsPopup, cmd_list.GetCommand(i));
			}
		}

		//コマンドのコントロールを [ツール] メニューに追加します:
		private void CreateCommand(ref object[] contextGUIDS, Commands2 commands, CommandBarPopup tools, CommandDef cmd)
		{
			if (tools == null)
				return;

			string commandName = cmd.Name;
			string commandText = cmd.Text;
			int commandIdx = cmd.Index;

			try
			{
				Command command = CreateCommand(ref contextGUIDS, commands, commandName, commandText, "");
				if (command == null)
					return;

				command.AddControl(tools.CommandBar, commandIdx);
			}
			catch (System.ArgumentException)
			{
				//同じ名前のコマンドが既に存在しているため、例外が発生した可能性があります。
				//  その場合、コマンドを再作成する必要はありません。 例外を 
				//  無視しても安全です。
			}
		}

		private Command CreateCommand(ref object[] contextGUIDS, Commands2 commands, string name, string buttonText, string tooltip)
		{
			return commands.AddNamedCommand2(_addInInstance, name, buttonText, tooltip, true, 59, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported + (int)vsCommandStatus.vsCommandStatusEnabled, (int)vsCommandStyle.vsCommandStylePictAndText, vsCommandControlType.vsCommandControlTypeButton);
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
				CommandList cmd_list = CreateOrGetCommandDef();
				if (GetCommandIndex(commandName, cmd_list) != -1)
				{
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
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
				if (HandleCommand(commandName))
				{
					handled = true;
					return;
				}
			}
		}

		bool HandleCommand(string commandName)
		{
			CommandList cmd_list = CreateOrGetCommandDef();
			int command_idx = GetCommandIndex(commandName, cmd_list);
			if (command_idx == -1)
				return false;

			CommandDef cmd_def = cmd_list.GetCommand(command_idx);
			if (cmd_def.Method != null)
			{
				cmd_def.Method(_applicationObject);
			}

			return true;
		}

		private CommandList CreateOrGetCommandDef()
		{
			if (_cmdList != null)
				return _cmdList;

			_cmdList = new CommandList();
			_cmdList.AddCommand( "SVN_Diff", (DTE2 vs_app) => (new TSvn( _applicationObject )).Diff() );
			_cmdList.AddCommand( "SVN_Log", (DTE2 vs_app) => (new TSvn( _applicationObject )).Log() );
			_cmdList.AddCommand( "GIT_Diff", (DTE2 vs_app) => (new TGit( _applicationObject )).Diff() );
			_cmdList.AddCommand( "GIT_Log", (DTE2 vs_app) => (new TGit( _applicationObject )).Log() );
			_cmdList.AddCommand( "Tortoise_Diff", (DTE2 vs_app) => (new TGit( _applicationObject )).Diff() );
			_cmdList.AddCommand( "Tortoise_Log", (DTE2 vs_app) => (new TGit( _applicationObject )).Log() );
			_cmdList.AddCommand( "Dummy", null );

			return _cmdList;
		}

		private int GetCommandIndex(string commandName, CommandList cmd_list)
		{
			string elem = commandName.Substring(_rootName.Length);
			for (int i = 0; i < cmd_list.NumCommands(); ++i)
			{
				if (elem == cmd_list.GetCommand(i).Name)
					return i;
			}
			return -1;
		}


		private DTE2 _applicationObject;
		private AddIn _addInInstance;

		private string _rootName = "TortoiseAddin.Connect.";

		CommandList _cmdList = null;
	}
}
