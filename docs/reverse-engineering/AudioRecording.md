# AudioRecording — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `AudioRecording` |
| XML `<type>` | `AudioRecording` |
| Parameter class | `AudioRecordingClusterParameter` |
| DLL class count | `AudioRecordingClusterParameter` (1 hit in LibConMas.dll) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `AudioRecordingClusterParameter`, `AudioRecording` (22 hits), `Audio` (29 hits broad) |
| **ConMas.iReporter.UserControls.dll** | `AudioRecordingParameter` class, `AudioRecordingPage`, `AudioRecordingPreparationPage`, `AudioReproducingPage`, `AudioRecordingWindow`, `RecordWindow`, `AudioRecordingFormat`, `AudioRecorder` |
| **ConMasClient.exe** | `AudioRecordingFileForm` |
| **LocalizableStrings.xml** | `Settings.Menu.VoiceRecognition`, `Settings.Detail.VoiceRecognition.*` (8 settings) |
| **NAudio.dll** | NAudio audio capture library (479KB) |
| **VoiceCore.dll** | Voice processing library (48KB) |

## Configuration Keys

### Verified Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether recording is required | Common parameter |
| `AudioRecordingFormat` | enum | (default) | WAV/MP3/WMA | Audio format | UserControls (`AudioRecordingFormat`, 1 hit) |
| `ReadOnly` | bool | 0 | 0/1 | Whether read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether hidden | UserControls |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `AudioMode` | enum | Recording mode (Normal/Voice) | VoiceRecognition settings |
| `AudioQuality` | enum | Recording quality | Common audio parameter |
| `AudioMaxDuration` | int | Maximum recording duration (s) | Common audio parameter |
| `AudioAutoStart` | bool | Auto-start recording | Pattern from other auto params |
| `AudioAutoStop` | bool | Auto-stop on silence | Pattern from other auto params |
| `AudioSampleRate` | int | Sample rate (Hz) | Common audio parameter |
| `AudioBitRate` | int | Bit rate (kbps) | Common audio parameter |
| `AudioChannels` | int | Channel count (1/2) | Common audio parameter |

## Runtime Controls (UserControls.dll)

### Recording Pages
- `AudioRecordingPage` — Main recording page
- `AudioRecordingPreparationPage` — Preparation/pre-recording page
- `AudioReproducingPage` — Playback/reproduction page
- `AudioRecordingPageViewModel` — MVVM view model

### Recording Windows
- `AudioRecordingWindow` — Audio recording window
- `RecordWindow` — Record popup window
- `CustomRecordWindow` — Custom record window
- `CustomMasterRecordWindow` — Master record window

### Recording UI Elements
- `RecordingStartImage` — Start recording icon
- `RecordingButtonImage` — Recording button icon
- `RecordingSeconds` — Recording time display
- `RecordingTime` — Recording time tracking

### Audio Libraries
- `NAudio.dll` — Audio capture and playback library (479KB)
- `VoiceCore.dll` — Voice processing (48KB)
- `MediaCore.dll` — Media core (11KB)

## Voice Recognition Settings (LocalizableStrings.xml)

Settings from `Settings.Detail.VoiceRecognition.*`:
- `UseVoiceRecognition` — Enable voice recognition
- `AllowsSpeechInputDuringAnswerBack` — Allow speech during answer-back
- `RevertToDefault` — Revert to default settings
- `VoicePitch` — Voice pitch (High/Low)
- `SpeakSpeed` — Speech speed (Fast/Slow)
- `EndDecisionTime` — End-of-speech detection (Rapid/Slow/Millisecond)

## Storage

- **Database:** `def_cluster.input_parameter` column
- **Data:** Audio file stored as BLOB/binary in `rep_cluster` (filename or binary data)

## Confidence Summary

| Item | Confidence |
|------|------------|
| AudioRecording as field type | ★★★★★ |
| AudioRecordingClusterParameter | ★★★★★ |
| Full recording UI (pages/windows) | ★★★★★ |
| NAudio capture library | ★★★★★ |
| AudioRecordingFormat parameter | ★★★☆☆ |
| Voice recognition settings | ★★★★☆ |
