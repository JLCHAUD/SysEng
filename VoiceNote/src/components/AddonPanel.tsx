import React, { useState } from 'react';
import { View, Text, TouchableOpacity, StyleSheet, ActivityIndicator } from 'react-native';
import { Ionicons } from '@expo/vector-icons';
import { getAddons } from '../addons';

interface Props {
  text: string;
  onApply: (newText: string) => void;
}

export default function AddonPanel({ text, onApply }: Props) {
  const [loading, setLoading] = useState<string | null>(null);
  const addons = getAddons();

  async function run(id: string, fn: (t: string) => Promise<string>) {
    setLoading(id);
    try {
      const result = await fn(text);
      onApply(result);
    } finally {
      setLoading(null);
    }
  }

  return (
    <View style={styles.container}>
      <Text style={styles.label}>Fonctions IA</Text>
      <View style={styles.row}>
        {addons.map((addon) => (
          <TouchableOpacity
            key={addon.id}
            style={styles.chip}
            onPress={() => run(addon.id, addon.run)}
            disabled={loading !== null}
            activeOpacity={0.7}
          >
            {loading === addon.id ? (
              <ActivityIndicator size="small" color="#6366f1" />
            ) : (
              <Ionicons name={addon.icon as any} size={16} color="#6366f1" />
            )}
            <Text style={styles.chipText}>{addon.name}</Text>
          </TouchableOpacity>
        ))}
      </View>
    </View>
  );
}

const styles = StyleSheet.create({
  container: { gap: 8 },
  label: { fontSize: 13, fontWeight: '600', color: '#64748b', textTransform: 'uppercase', letterSpacing: 0.8 },
  row: { flexDirection: 'row', flexWrap: 'wrap', gap: 8 },
  chip: {
    flexDirection: 'row', alignItems: 'center', gap: 6,
    paddingHorizontal: 12, paddingVertical: 8,
    backgroundColor: '#eef2ff', borderRadius: 20,
    borderWidth: 1, borderColor: '#c7d2fe',
  },
  chipText: { fontSize: 13, color: '#4f46e5', fontWeight: '500' },
});
