import * as React from 'react';
import { useState } from 'react';
import { v4 as uuidv4 } from 'uuid';
import {
    Stack,
    Text,
    PrimaryButton,
    Image,
    ImageFit,
    TextField,
    Label,
    MessageBar,
    MessageBarType,
    IStackTokens,
    IImageProps,
    IconButton,
    IIconProps
} from '@fluentui/react';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import { IImageUploadProps, IQuizImage } from './interfaces';
import styles from './Quiz.module.scss';

// Icons
const uploadIcon: IIconProps = { iconName: 'Upload' };
const removeIcon: IIconProps = { iconName: 'Delete' };
const editIcon: IIconProps = { iconName: 'Edit' };

// Stack tokens
const stackTokens: IStackTokens = {
    childrenGap: 10
};

// Default props
const DEFAULT_MAX_SIZE_MB = 5;
const DEFAULT_ACCEPT = "image/png,image/jpeg,image/gif";

const ImageUpload: React.FC<IImageUploadProps> = (props) => {
    const {
        onImageUpload,
        onImageRemove,
        currentImage,
        label = "Upload Image",
        accept = DEFAULT_ACCEPT,
        maxSizeMB = DEFAULT_MAX_SIZE_MB,
        context
    } = props;

    const [image, setImage] = useState<IQuizImage | undefined>(currentImage);
    const [error, setError] = useState<string>('');
    const [altText, setAltText] = useState<string>(currentImage?.altText || '');
    const [showFilePicker, setShowFilePicker] = useState<boolean>(false);

    // Handle file picker selection
    const handleFilePickerSave = (filePickerResults: IFilePickerResult[]): void => {


        // Use the first file if multiple were selected
        const filePickerResult = filePickerResults && filePickerResults.length > 0 ? filePickerResults[0] : null;

        if (!filePickerResult) {
            setError('No file was selected');
            return;
        }

        // Validate file size
        filePickerResult.downloadFileContent().then((file: File) => {
            if (file.size > maxSizeMB * 1024 * 1024) {
                setError(`File size exceeds the maximum allowed size of ${maxSizeMB}MB.`);
                return;
            }

            // Create image object
            const newImage: IQuizImage = {
                id: uuidv4(),
                url: filePickerResult.fileAbsoluteUrl,
                fileName: filePickerResult.fileName,
                altText: altText || filePickerResult.fileNameWithoutExtension
            };

            setImage(newImage);
            setAltText(newImage.altText || '');
            onImageUpload(newImage);
            setError('');
        }).catch((err: Error) => {
            setError(`Error processing file: ${err.message}`);
        });


        setShowFilePicker(false);
    };

    // Handle file picker cancel
    const handleFilePickerCancel = (): void => {
        setShowFilePicker(false);
    };

    // Handle alt text change
    const handleAltTextChange = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        setAltText(newValue || '');

        if (image) {
            const updatedImage: IQuizImage = {
                ...image,
                altText: newValue || ''
            };
            setImage(updatedImage);
            onImageUpload(updatedImage);
        }
    };

    // Handle image removal
    const handleRemoveImage = (): void => {
        setImage(undefined);
        setAltText('');

        if (onImageRemove) {
            onImageRemove();
        }
    };

    // Image component props
    const imageProps: IImageProps = {
        src: image?.url,
        alt: image?.altText || 'Uploaded image',
        imageFit: ImageFit.contain,
        maximizeFrame: true,
        width: '100%',
        height: 200
    };

    return (
        <div className={styles.imageUploadContainer}>
            <Label>{label}</Label>

            {error && (
                <MessageBar
                    messageBarType={MessageBarType.error}
                    isMultiline={false}
                    onDismiss={() => setError('')}
                    dismissButtonAriaLabel="Close"
                >
                    {error}
                </MessageBar>
            )}

            {image ? (
                // Display uploaded image with edit controls
                <Stack tokens={stackTokens}>
                    <div className={styles.imagePreview}>
                        <Image {...imageProps} />
                        <div className={styles.imageControls}>
                            <IconButton
                                iconProps={editIcon}
                                title="Change Image"
                                ariaLabel="Change Image"
                                onClick={() => setShowFilePicker(true)}
                            />
                            <IconButton
                                iconProps={removeIcon}
                                title="Remove Image"
                                ariaLabel="Remove Image"
                                onClick={handleRemoveImage}
                            />
                        </div>
                    </div>

                    <TextField
                        label="Alt Text (for accessibility)"
                        value={altText}
                        onChange={handleAltTextChange}
                        placeholder="Describe the image"
                    />

                    <Text variant="small" className={styles.imageMetadata}>
                        Filename: {image.fileName}
                    </Text>
                </Stack>
            ) : (
                // Display upload button when no image is present
                <Stack tokens={stackTokens} horizontalAlign="center">
                    <div className={styles.uploadPlaceholder}>
                        <PrimaryButton
                            iconProps={uploadIcon}
                            text="Select Image"
                            onClick={() => setShowFilePicker(true)}
                        />
                        <Text variant="small">
                            Supported formats: PNG, JPEG, GIF (max {maxSizeMB}MB)
                        </Text>
                    </div>
                </Stack>
            )}

            {showFilePicker && context && (
                <FilePicker
                    context={context}
                    accepts={[accept]} // Now accepting array format
                    buttonLabel="Select Image"
                    hideWebSearchTab={true}
                    onSave={handleFilePickerSave}
                    onCancel={handleFilePickerCancel}
                    key={Date.now()} // Force re-render each time
                />
            )}
        </div>
    );
};

export default ImageUpload;