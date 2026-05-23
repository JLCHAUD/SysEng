import 'react-native-get-random-values';
import React, { useState } from 'react';
import { StatusBar } from 'expo-status-bar';
import HomeScreen from './src/screens/HomeScreen';
import RecordScreen from './src/screens/RecordScreen';
import DetailScreen from './src/screens/DetailScreen';
import { Recording } from './src/types';

type Screen = { name: 'home' } | { name: 'record' } | { name: 'detail'; recording: Recording };

export default function App() {
  const [screen, setScreen] = useState<Screen>({ name: 'home' });

  if (screen.name === 'record') {
    return (
      <>
        <StatusBar style="dark" />
        <RecordScreen
          onDone={(rec) => setScreen({ name: 'detail', recording: rec })}
          onBack={() => setScreen({ name: 'home' })}
        />
      </>
    );
  }

  if (screen.name === 'detail') {
    return (
      <>
        <StatusBar style="dark" />
        <DetailScreen
          recording={screen.recording}
          onBack={() => setScreen({ name: 'home' })}
        />
      </>
    );
  }

  return (
    <>
      <StatusBar style="dark" />
      <HomeScreen
        onOpenRecording={(rec) => setScreen({ name: 'detail', recording: rec })}
        onNewRecording={() => setScreen({ name: 'record' })}
      />
    </>
  );
}
