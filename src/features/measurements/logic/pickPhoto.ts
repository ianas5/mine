import * as ImagePicker from 'expo-image-picker';

export interface PickedImage {
  readonly uri: string;
  readonly width: number | null;
  readonly height: number | null;
}

const toPicked = (result: ImagePicker.ImagePickerResult): PickedImage | null => {
  const asset = result.canceled ? undefined : result.assets[0];
  if (!asset) return null;
  return { uri: asset.uri, width: asset.width || null, height: asset.height || null };
};

/** Pick a progress photo from the library (no permission prompt on the system picker). */
export async function pickFromLibrary(): Promise<PickedImage | null> {
  return toPicked(
    await ImagePicker.launchImageLibraryAsync({ mediaTypes: 'images', quality: 0.85 }),
  );
}

/** Capture a progress photo with the camera (requests camera permission first). */
export async function captureFromCamera(): Promise<PickedImage | null> {
  const permission = await ImagePicker.requestCameraPermissionsAsync();
  if (!permission.granted) return null;
  return toPicked(await ImagePicker.launchCameraAsync({ mediaTypes: 'images', quality: 0.85 }));
}
