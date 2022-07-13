﻿/*
 * GSClient.cs
 *
 * authors: Tomasz Rybiarczyk (tomasz.rybiarczyk@steelseries.com)
 *
 *
 * Copyright (c) 2016 SteelSeries
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
 * IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
 * CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
 * TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
 * SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

using System.IO;
using System.Collections.Generic;
using System.Text.Json;


namespace SteelSeries
{

    namespace GameSense
    {

        public class GSClient
        {

            // ******************** EDITOR VISIBLE  ********************

            public string GameName;
            public string GameDisplayName;
            public string Developer;
            public Bind_Event[] Events;




            // ******************** TYPES ********************

            [System.Serializable]
            struct coreProps
            {
#pragma warning disable 0649
                public string address;
#pragma warning restore 0649
            }

            enum ClientState
            {
                Active,
                Probing,
                Inactive
            }

            [FullSerializer.fsObject(Converter = typeof(BindEventConverter))]
            [System.Serializable]
            sealed public class Bind_Event
            {

                [System.NonSerialized] public string game;
                public string eventName;
                public System.Int32 minValue;
                public System.Int32 maxValue;
                public EventIconId iconId;
                public bool valueOptional;
                public AbstractHandler[] handlers;

                public Bind_Event(string gameName, string eventName, System.Int32 minValue, System.Int32 maxValue, EventIconId iconId, AbstractHandler[] handlers, bool valueOptional = false)
                {
                    this.game = gameName;
                    this.eventName = eventName;
                    this.minValue = minValue;
                    this.maxValue = maxValue;
                    this.iconId = iconId;
                    this.handlers = handlers;
                    this.valueOptional = valueOptional;
                }
            }

            [FullSerializer.fsObject(Converter = typeof(RegisterEventConverter))]
            sealed class Register_Event
            {

                [System.NonSerialized] public string game;
                public string eventName;
                public System.Int32 minValue;
                public System.Int32 maxValue;
                public EventIconId iconId;
                public bool valueOptional;

                public Register_Event(string gameName, string eventName, System.Int32 minValue, System.Int32 maxValue, EventIconId iconId, bool valueOptional = false)
                {
                    this.game = gameName;
                    this.eventName = eventName;
                    this.minValue = minValue;
                    this.maxValue = maxValue;
                    this.iconId = iconId;
                    this.valueOptional = valueOptional;
                }
            }

            [FullSerializer.fsObject(Converter = typeof(SendEventConverter))]
            sealed class Send_Event
            {
                public string game;
                public string event_name;
                public EventData data;
            }

            [FullSerializer.fsObject(Converter = typeof(RegisterGameConverter))]
            sealed class Register_Game
            {
                public string game;
                public string game_display_name;
                public string developer;
            }

            sealed class Game
            {
                public string game;
                public Game(string gameName)
                {
                    game = gameName;
                }
            }


            class BindEventConverter : Converter<Bind_Event>
            {
                protected override FullSerializer.fsResult DoSerialize(Bind_Event model, System.Collections.Generic.Dictionary<string, FullSerializer.fsData> serialized)
                {

                    SerializeMember(serialized, null, "game", model.game.ToUpper());
                    SerializeMember(serialized, null, "event", model.eventName);
                    SerializeMember(serialized, null, "min_value", model.minValue);
                    SerializeMember(serialized, null, "max_value", model.maxValue);
                    SerializeMember(serialized, null, "icon_id", (System.UInt32)model.iconId);
                    SerializeMember(serialized, null, "value_optional", model.valueOptional);
                    SerializeMember(serialized, null, "handlers", model.handlers);

                    return FullSerializer.fsResult.Success;
                }
            }

            class RegisterEventConverter : Converter<Register_Event>
            {
                protected override FullSerializer.fsResult DoSerialize(Register_Event model, System.Collections.Generic.Dictionary<string, FullSerializer.fsData> serialized)
                {

                    SerializeMember(serialized, null, "game", model.game.ToUpper());
                    SerializeMember(serialized, null, "event", model.eventName);
                    SerializeMember(serialized, null, "min_value", model.minValue);
                    SerializeMember(serialized, null, "max_value", model.maxValue);
                    SerializeMember(serialized, null, "icon_id", (System.UInt32)model.iconId);
                    SerializeMember(serialized, null, "value_optional", model.valueOptional);

                    return FullSerializer.fsResult.Success;
                }
            }

            class SendEventConverter : Converter<Send_Event>
            {
                protected override FullSerializer.fsResult DoSerialize(Send_Event model, System.Collections.Generic.Dictionary<string, FullSerializer.fsData> serialized)
                {

                    SerializeMember<string>(serialized, null, "game", model.game);
                    SerializeMember<string>(serialized, null, "event", model.event_name);
                    SerializeMember<EventData>(serialized, null, "data", model.data);

                    return FullSerializer.fsResult.Success;
                }
            }

            class RegisterGameConverter : Converter<Register_Game>
            {
                protected override FullSerializer.fsResult DoSerialize(Register_Game model, System.Collections.Generic.Dictionary<string, FullSerializer.fsData> serialized)
                {

                    SerializeMember(serialized, null, "game", model.game);
                    SerializeMember(serialized, null, "game_display_name", model.game_display_name);
                    SerializeMember(serialized, null, "developer", model.developer);

                    return FullSerializer.fsResult.Success;
                }
            }


            abstract class AbstractQueueMsg
            {
                public abstract System.Uri uri { get; }
                public abstract bool IsCritical();
                public abstract string ToJson(FullSerializer.fsSerializer serializer);
            }

            abstract class QueueMsg<T> : AbstractQueueMsg
            {
                public T data { get; set; }
                public override string ToJson(FullSerializer.fsSerializer serializer)
                {
                    string serialized = null;
                    FullSerializer.fsData fsData;
                    FullSerializer.fsResult fsResult;

                    fsResult = serializer.TrySerialize<T>(data, out fsData);

                    if (fsResult.Succeeded)
                    {
#if SS_DEBUG
                        serialized = FullSerializer.fsJsonPrinter.PrettyJson( fsData );
#else
                        serialized = FullSerializer.fsJsonPrinter.CompressedJson(fsData);
#endif
                    }
                    else
                    {
                        throw new System.Exception("Failed serializing object: " + data.ToString());
                    }

                    return serialized;
                }
            }

            class QueueMsgRegisterGame : QueueMsg<Register_Game>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return true; }
            }

            class QueueMsgBindEvent : QueueMsg<Bind_Event>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return true; }
            }

            class QueueMsgRegisterEvent : QueueMsg<Register_Event>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return true; }
            }

            class QueueMsgSendEvent : QueueMsg<Send_Event>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return false; }
            }

            class QueueMsgSendHeartbeat : QueueMsg<Game>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return false; }
            }

            class QueueMsgRemoveGame : QueueMsg<Game>
            {
                public static System.Uri _uri;
                public override System.Uri uri { get { return _uri; } }
                public override bool IsCritical() { return false; }
            }


            class CriticalMessageIllFormedException : System.Exception
            {
                public CriticalMessageIllFormedException(string message) : base(message) { }
            }

            class ServerDownException : System.Exception
            {
                public ServerDownException(string message) : base(message) { }
            }


            // ******************** CONSTANTS ********************

            private const string _SceneObjName = "GameSenseManager_Auto";
            private const string _GameSenseObjName = "GameSenseManager";
            private const uint _MsgQueueSize = 100;
            private const int _ServerProbeInterval = 5000; // 5 seconds
            private const int _MsgCheckInterval = 10;   // 10 ms
            private const long _MaxIdleTimeBeforeHeartbeat = 1000L;   // 1 second
            private const string _ServerPropsPath = "%PROGRAMDATA%/SteelSeries/SteelSeries Engine 3/coreProps.json";

            // ******************** DATA MEMBERS ********************

            private static GSClient _mInstance;
            private ClientState _mClientState;
            private string _mServerPort;

            private System.Threading.Thread _mGameSenseWrk;
            private bool _mGameSenseWrkShouldRun;
            private LocklessQueue<AbstractQueueMsg> _mMsgQueue;

            private FullSerializer.fsSerializer _mSerializer;

            private System.Uri _uriBase;

            private HashSet<string> _mEventList;

            // ******************** PROPS ********************

            public static GSClient Instance
            {
                get
                {
                    if (_mInstance == null)
                    {
                        _mInstance = new();
                    }
                    return _mInstance;
                }
            }




            // ******************** METHODS ********************

            private static void _logException(string msg, System.Exception e)
            {
#if !SS_NWARN
                System.Console.WriteLine( $"[GameSense Client] {msg} - {e.Message}\n{e.StackTrace}" );
#endif
            }


            private static void _logWarning(string msg)
            {
#if !SS_NWARN
                System.Console.WriteLine( $"[GameSense Client] {msg}" );
#endif
            }


            private static void _logDbgMsg(string msg)
            {
#if SS_DEBUG
                System.Console.WriteLine( $"[GameSense Client] {msg}" );
#endif
            }


            private static void _logErrorMsg(string msg)
            {
                System.Console.WriteLine( $"[GameSense Client] {msg}" );
            }


            private void _setClientState(ClientState state)
            {
                _mClientState = state;
            }


            private bool _isClientActiveOrProbing()
            {
                return _mClientState == ClientState.Active || _mClientState == ClientState.Probing;
            }


            private void _sendServer(System.Uri uri, string data)
            {
                byte[] payload = System.Text.Encoding.ASCII.GetBytes(data);

                System.Net.HttpWebRequest req = (System.Net.HttpWebRequest)System.Net.HttpWebRequest.Create(uri);
                req.ContentType = "application/json";
                req.Method = "POST";

                Stream strm = req.GetRequestStream();
                strm.Write(payload, 0, payload.Length);
                strm.Close();

                System.Net.HttpWebResponse resp = (System.Net.HttpWebResponse)req.GetResponse();
#if SS_DEBUG
                strm = resp.GetResponseStream();
                StreamReader sr = new StreamReader( strm, System.Text.Encoding.UTF8 );
                _logDbgMsg( "Server response: " + sr.ReadToEnd() );
                strm.Close();
#endif
                resp.Close();
            }


            private static string _getPropsPath()
            {
                return System.Environment.ExpandEnvironmentVariables( _ServerPropsPath );
            }


            private static string _readProps()
            {
                string propsPath = _getPropsPath();
                string propsJson = null;

                try
                {
                    propsJson = System.IO.File.ReadAllText(propsPath);
                }
                catch (System.Exception e)
                {
                    _logException("Could not read server props file", e);
                }

                return propsJson;
            }


            private static string _getServerPort()
            {
                string port = null;
                string propsJson = _readProps();

                if (propsJson != null)
                {
                    try
                    {
                        // { address: "127.0.0.1:xxxxx" }
                        var jsonDoc = JsonDocument.Parse(propsJson);
                        var address = jsonDoc.RootElement.GetProperty("address").GetString();
                        string[] addrPort = address.Split(':');

                        // this will raise an exception if our input is not what we expect
                        System.Convert.ToUInt16(addrPort[1]);

                        port = addrPort[1];
                    }
                    catch (System.Exception e)
                    {
                        _logException("Cannot parse port information", e);
                    }
                }

                return port;
            }


            private void _initializeUris()
            {
                _uriBase = new System.Uri("http://127.0.0.1:" + _mServerPort);

                // register uris with corresponding msg types
                QueueMsgRegisterGame._uri = new System.Uri(_uriBase, "game_metadata");
                QueueMsgBindEvent._uri = new System.Uri(_uriBase, "bind_game_event");
                QueueMsgRegisterEvent._uri = new System.Uri(_uriBase, "register_game_event");
                QueueMsgSendEvent._uri = new System.Uri(_uriBase, "game_event");
                QueueMsgSendHeartbeat._uri = new System.Uri(_uriBase, "game_heartbeat");
                QueueMsgRemoveGame._uri = new System.Uri(_uriBase, "remove_game");
            }


            // parse and send msg
            // will raise an exception on any failure
            private void _sendMsg(AbstractQueueMsg msg)
            {
                string data = msg.ToJson(_mSerializer);
                _logDbgMsg(data);

                try
                {
                    _sendServer(msg.uri, data);
                }
                catch (System.Net.WebException e)
                {

                    switch (e.Status)
                    {

                        // retrieve syntax error message from the server
                        case System.Net.WebExceptionStatus.ProtocolError:
                            if (msg.IsCritical())
                            {
                                System.Net.WebResponse resp = e.Response;
                                Stream strm = resp.GetResponseStream();
                                StreamReader sr = new StreamReader(strm, System.Text.Encoding.UTF8);
                                string serverMsg = sr.ReadToEnd();
                                strm.Close();
                                throw new CriticalMessageIllFormedException(serverMsg);
                            }
                            break;

                        case System.Net.WebExceptionStatus.ConnectFailure:
                            throw new ServerDownException(e.Message);

                        // unhandled
                        default:
                            _logException("Unexpected status", e);
                            throw;

                    }
                }
            }


            private void _addGUIDefinedEvents()
            {

                if (GameName == null || GameDisplayName == null || Events == null)
                {
                    _logWarning("Incomplete game registration form");
                    _setClientState(ClientState.Inactive);
                    return;
                }

                RegisterGame(GameName, GameDisplayName, Developer);

                // TODO need to throttle this so we do not overflow the queue
                //      temporarily mitigated by increasing the queue size
                foreach (Bind_Event be in Events)
                {
                    _mEventList.Add(be.eventName);
                    AbstractQueueMsg msg;

                    if (be.handlers == null || be.handlers.Length == 0)
                    {

                        // if no handlers provided, we will at least register these events with SSE
                        msg = new QueueMsgRegisterEvent
                        {
                            data = new Register_Event(GameName, be.eventName, be.minValue, be.maxValue, be.iconId)
                        };

                    }
                    else
                    {

                        // bind events with defualt handlers
                        be.game = GameName;
                        msg = new QueueMsgBindEvent
                        {
                            data = be
                        };

                        // let handlers do work
                        // on main thread before enqueueing
                        foreach (var h in be.handlers)
                        {
                            h.Preprocess();
                        }

                    }

                    _mMsgQueue.PEnqueue(msg);
                }
            }


            private void _gamesenseWrk()
            {
                AbstractQueueMsg pendingMsg = null;
                System.Diagnostics.Stopwatch tLastMsg = new System.Diagnostics.Stopwatch();
                tLastMsg.Start();

                while (_mGameSenseWrkShouldRun)
                {


                    switch (_mClientState)
                    {

                        case ClientState.Active:

                            // see if there is any message to process
                            AbstractQueueMsg msg;
                            while ((msg = _mMsgQueue.CDequeue()) == null && _mGameSenseWrkShouldRun)
                            {
                                // no messages in queue, sleep a bit before checking again
                                System.Threading.Thread.Sleep(_MsgCheckInterval);

                                // if a heartbeat event is due, send it now
                                if (tLastMsg.ElapsedMilliseconds > _MaxIdleTimeBeforeHeartbeat)
                                {
                                    msg = new QueueMsgSendHeartbeat
                                    {
                                        data = new Game(GameName)
                                    };

                                    break;
                                }
                            }

                            // Make sure the top-level state is still active before attempting to send a message
                            if (!_mGameSenseWrkShouldRun)
                            {
                                break;
                            }

                            try
                            {

                                // attempt to send a message
                                _sendMsg(msg);

                                // message sent successfully, reset timer
                                tLastMsg.Reset();
                                tLastMsg.Start();

                            }
                            catch (ServerDownException e)
                            {

                                // GameSense server down, save the message and go to probing state
                                _logException("Failed connecting to GameSense server", e);
                                pendingMsg = msg;
                                _setClientState(ClientState.Probing);

                            }
                            catch (CriticalMessageIllFormedException e)
                            {

                                // RegisterGame or RegisterEvent or BindEvent messages are not well formed
                                // log and disable
                                _logException("Message ill-formed", e);
                                _setClientState(ClientState.Inactive);

                            }
                            catch (System.Exception e)
                            {
                                _logException("Failed processing msg", e);
                            }

                            break;

                        case ClientState.Probing:

                            // parse props to get GameSense server port
                            _mServerPort = _getServerPort();
                            if (_mServerPort == null)
                            {
                                // SSE3 not installed or coreprops garbled
                                // this failure is beyond anything we can do, GameSense will remain disabled
                                _logWarning("Failed to obtain GameSense server port. GameSense will not function");
                                _setClientState(ClientState.Inactive);
                                break;
                            }

                            // port information successfully parsed, init uris
                            _initializeUris();

                            if (pendingMsg != null)
                            {

                                // attempt to send a pending message
                                try
                                {
                                    _sendMsg(pendingMsg);
                                    pendingMsg = null;
                                }
                                catch (ServerDownException e)
                                {
                                    // GameSense server down, wait a bit and retry
                                    _logException("Failed connecting to GameSense server", e);
                                    _logDbgMsg("Retrying in 5 seconds...");
                                    System.Threading.Thread.Sleep(_ServerProbeInterval);
                                    break;
                                }

                            }

                            // we can attempt to send data, go to active state
                            _setClientState(ClientState.Active);
                            break;

                        case ClientState.Inactive:
                            // quit
                            _logDbgMsg("Entering inactive state");
                            _mGameSenseWrkShouldRun = false;
                            break;

                        default:
                            // unknown state, must abort
                            _logErrorMsg("Unknown GameSense client state");
                            _setClientState(ClientState.Inactive);
                            break;

                    }


                }

                _logDbgMsg("Worker exiting");

            }


            // Initialization
            public void Initialize()
            {

#if !SS_GAMESENSE_DISABLED
                // initialize
                _mSerializer = new FullSerializer.fsSerializer();
                _mMsgQueue = new LocklessQueue<AbstractQueueMsg>(_MsgQueueSize);
                _mGameSenseWrk = new System.Threading.Thread(_gamesenseWrk); // check for exceptions
                _mGameSenseWrkShouldRun = true;
                _setClientState(ClientState.Probing);
                _mEventList = new HashSet<string>();

                try
                {
                    _mGameSenseWrk.Start();
                }
                catch (System.Exception e)
                {
                    _logException("Could not start the client thread", e);
                    _setClientState(ClientState.Inactive);
                }

#endif  // !SS_GAMESENSE_DISABLED

            }

            public void Terminate()
            {
                _setClientState(ClientState.Inactive);
            }

            // ******************** API ********************

            /// <summary>
            /// Registers a game with GameSense server. Must be called before everything else.
            /// </summary>
            /// <param name="name">Game Identifier. Valid characters are limited to uppercase A-Z, the digits 0-9, hyphen, and underscore</param>
            /// <param name="displayName">The name that will appear in SteelSeries Engine for this game</param>
            /// <param name="iconColor">Color icon identifier that will appear for this game in SteelSeries Engine</param>
            [System.Obsolete("This method is deprecated. See RegisterGame below.")]
            public void RegisterGame(string name, string displayName, IconColor iconColor)
            {
#if !SS_GAMESENSE_DISABLED
                GameName = name.ToUpper();
                GameDisplayName = displayName;
                Developer = "";

                Register_Game obj = new Register_Game
                {
                    game = GameName,
                    game_display_name = GameDisplayName,
                    developer = Developer
                };

                QueueMsgRegisterGame msg = new QueueMsgRegisterGame
                {
                    data = obj
                };

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

            /// <summary>
            /// Registers a game with GameSense server. Must be called before everything else.
            /// </summary>
            /// <param name="name">Game Identifier. Valid characters are limited to uppercase A-Z, the digits 0-9, hyphen, and underscore</param>
            /// <param name="displayName">The name that will appear in SteelSeries Engine for this game</param>
            /// <param name="iconColor">Color icon identifier that will appear for this game in SteelSeries Engine</param>
            public void RegisterGame(string name, string displayName, string developer)
            {
#if !SS_GAMESENSE_DISABLED
                GameName = name.ToUpper();
                GameDisplayName = displayName;
                Developer = developer;

                Register_Game obj = new Register_Game
                {
                    game = GameName,
                    game_display_name = GameDisplayName,
                    developer = Developer
                };

                QueueMsgRegisterGame msg = new QueueMsgRegisterGame
                {
                    data = obj
                };

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

            /// <summary>
            /// Defines an event with its handlers and registers them with the GameSense Server.
            /// </summary>
            /// <param name="eventName">Event name. Valid characters are limited to uppercase A-Z, the digits 0-9, hyphen, and underscore</param>
            /// <remarks>
            /// You should not call this function more than once for the same event. Doing so will unbind previously bound handlers.
            /// </remarks>
            /// <param name="minValue">Minimum value</param>
            /// <param name="maxValue">Maximum value</param>
            /// <param name="iconId">Identifies an icon that will be show in SteelSeries Engine for this event</param>
            /// <param name="handlers">An array of handlers for this event</param>
            /// <param name="valueOptional">f true, handlers will be processed on every event update with no regard to update value</param>
            public void BindEvent(string eventName, System.Int32 minValue, System.Int32 maxValue, EventIconId iconId, AbstractHandler[] handlers, bool valueOptional = false)
            {
#if !SS_GAMESENSE_DISABLED
                if (!_isClientActiveOrProbing()) return;
                _mEventList.Add(eventName);

                QueueMsgBindEvent msg = new QueueMsgBindEvent
                {
                    data = new Bind_Event(GameName.ToUpper(), eventName.ToUpper(), minValue, maxValue, iconId, handlers, valueOptional)
                };

                // let handlers do work
                // on main thread before enqueueing
                foreach (var h in handlers)
                {
                    h.Preprocess();
                }

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

            /// <summary>
            /// Defines an event and registers it with the GameSense Server.
            /// </summary>
            /// <param name="eventName">Event name. Valid characters are limited to uppercase A-Z, the digits 0-9, hyphen, and underscore</param>
            /// <remarks>
            /// You should call this function if you do not want specify default behavior for events but rather want it exposed in SteelSeries Engine for the user to customize.
            /// If you wish to specify default behavior (handlers), use <see cref="GSClient.BindEvent"/> instead
            /// </remarks>
            /// <param name="minValue">Minimum value</param>
            /// <param name="maxValue">Maximum value</param>
            /// <param name="iconId">Identifies an icon that will be show in SteelSeries Engine for this event</param>
            /// <param name="valueOptional">f true, handlers will be processed on every event update with no regard to update value</param>
            public void RegisterEvent(string eventName, System.Int32 minValue, System.Int32 maxValue, EventIconId iconId, bool valueOptional = false)
            {
#if !SS_GAMESENSE_DISABLED
                if (!_isClientActiveOrProbing()) return;
                _mEventList.Add(eventName);

                QueueMsgRegisterEvent msg = new QueueMsgRegisterEvent
                {
                    data = new Register_Event(GameName.ToUpper(), eventName.ToUpper(), minValue, maxValue, iconId, valueOptional)
                };

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

            /// <summary>
            /// Sends a new value for the specified event to the GameSense Server
            /// </summary>
            /// <param name="eventName">Previously bound/registered event</param>
            /// <param name="value">New value</param>
            /// <param name="frame">Optional context frame specifying additional</param>
            public void SendEvent(string eventName, System.Int32 value, AbstractContextFrame frame = null)
            {
#if !SS_GAMESENSE_DISABLED
                if (!_isClientActiveOrProbing()) return;
                if (!_mEventList.Contains(eventName))
                {
                    _logWarning("Attempting to send unregistered event: " + eventName);
                }

                Send_Event se = new Send_Event
                {
                    game = GameName,
                    event_name = eventName.ToUpper(),
                    data = {
                        value = value,
                        frame = frame
                    }
                };

                QueueMsgSendEvent msg = new QueueMsgSendEvent
                {
                    data = se
                };

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

            /// <summary>
            /// Notifies the GameSense Server to remove the game and all registered events
            /// </summary>
            public void RemoveGame()
            {
#if !SS_GAMESENSE_DISABLED
                if (!_isClientActiveOrProbing()) return;

                QueueMsgRemoveGame msg = new QueueMsgRemoveGame
                {
                    data = new Game(GameName)
                };

                _mMsgQueue.PEnqueue(msg);
#endif  // !SS_GAMESENSE_DISABLED
            }

        }
    }
}




class LocklessQueue<T>
{

    private T[] _bfr;
    private System.Int32 _length;
    private System.Int32 _readIdx;
    private System.Int32 _maxReadIdx;
    private System.Int32 _writeIdx;


    public LocklessQueue(uint size)
    {
        _bfr = new T[size];
        _length = (System.Int32)size;
        _readIdx = 0;
        _maxReadIdx = 0;
        _writeIdx = 0;
    }

    private System.Int32 index(System.Int32 i)
    {
        return i % _length;
    }

    public bool PEnqueue(T obj)
    {
        System.Int32 readIdx;
        System.Int32 writeIdx;

        do
        {
            // TODO beware of overflow
            writeIdx = _writeIdx;
            readIdx = _readIdx;

            if (index(writeIdx + 1) == index(readIdx))
            {
                return false;
            }

            // repeat if _writeIdx has mutated
        } while (System.Threading.Interlocked.CompareExchange(ref _writeIdx, index(writeIdx + 1), writeIdx) != writeIdx);

        _bfr[index(writeIdx)] = obj;

        // commit
        while (System.Threading.Interlocked.CompareExchange(ref _maxReadIdx, index(writeIdx + 1), writeIdx) != writeIdx)
        {
            // TODO yield? naaah!
        }

        return true;
    }

    public T CDequeue()
    {
        T obj = default(T);
        System.Int32 maxReadIdx;
        System.Int32 readIdx;

        do
        {

            readIdx = _readIdx;
            maxReadIdx = _maxReadIdx;

            if (index(readIdx) == index(maxReadIdx))
            {
                return default(T);
            }

            obj = _bfr[index(readIdx)];

            if (System.Threading.Interlocked.CompareExchange(ref _readIdx, (readIdx + 1), readIdx) != _readIdx)
            {
                break;
            }

        } while (true);

        return obj;
    }
}
