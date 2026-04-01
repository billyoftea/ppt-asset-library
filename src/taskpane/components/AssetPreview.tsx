import React, { useCallback, useState } from "react";
import {
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Text,
  Badge,
  Spinner,
  useToastController,
  Toast,
  ToastBody,
  ToastTitle,
} from "@fluentui/react-components";
import {
  DismissRegular,
  SlideAddRegular,
  ArrowLeftRegular,
  ArrowRightRegular,
  ImageRegular,
} from "@fluentui/react-icons";
import { Asset } from "../../types";
import { insertSlide } from "../../services/officeApi";

interface AssetPreviewProps {
  asset: Asset;
  onClose: () => void;
  toasterId: string;
  /** All currently visible assets for prev/next navigation */
  allAssets?: Asset[];
  /** Navigate to a different asset */
  onNavigate?: (asset: Asset) => void;
}

const AssetPreview: React.FC<AssetPreviewProps> = ({
  asset,
  onClose,
  toasterId,
  allAssets,
  onNavigate,
}) => {
  const { dispatchToast } = useToastController(toasterId);
  const [inserting, setInserting] = useState(false);
  const [imgError, setImgError] = useState(false);

  // Keyboard navigation
  const currentIndex = allAssets?.findIndex((a) => a.id === asset.id) ?? -1;
  const hasPrev = currentIndex > 0;
  const hasNext = allAssets ? currentIndex < allAssets.length - 1 : false;

  const navigatePrev = useCallback(() => {
    if (hasPrev && allAssets && onNavigate) {
      setImgError(false);
      onNavigate(allAssets[currentIndex - 1]);
    }
  }, [hasPrev, allAssets, onNavigate, currentIndex]);

  const navigateNext = useCallback(() => {
    if (hasNext && allAssets && onNavigate) {
      setImgError(false);
      onNavigate(allAssets[currentIndex + 1]);
    }
  }, [hasNext, allAssets, onNavigate, currentIndex]);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === "ArrowLeft") {
        e.preventDefault();
        navigatePrev();
      } else if (e.key === "ArrowRight") {
        e.preventDefault();
        navigateNext();
      }
    },
    [navigatePrev, navigateNext]
  );

  const handleInsertSlide = useCallback(async () => {
    setInserting(true);
    try {
      await insertSlide(asset);
      dispatchToast(
        <Toast>
          <ToastTitle>Inserted!</ToastTitle>
          <ToastBody>
            Slide #{asset.slideNumber} has been inserted as a fully editable native slide.
          </ToastBody>
        </Toast>,
        { intent: "success" }
      );
      onClose();
    } catch (err) {
      dispatchToast(
        <Toast>
          <ToastTitle>Insert failed</ToastTitle>
          <ToastBody>
            {err instanceof Error ? err.message : "Unknown error"}
          </ToastBody>
        </Toast>,
        { intent: "error" }
      );
    } finally {
      setInserting(false);
    }
  }, [asset, dispatchToast, onClose]);

  return (
    <Dialog open={true} onOpenChange={(_, data) => !data.open && onClose()}>
      <DialogSurface className="preview-surface" onKeyDown={handleKeyDown}>
        <DialogBody>
          <DialogTitle
            action={
              <Button
                appearance="subtle"
                icon={<DismissRegular />}
                onClick={onClose}
                aria-label="Close"
              />
            }
          >
            <div className="preview-title-row">
              <Text weight="semibold" size={400} className="preview-title-text">
                Slide #{asset.slideNumber}
              </Text>
              {allAssets && allAssets.length > 1 && (
                <Text size={100} className="preview-counter">
                  {currentIndex + 1} / {allAssets.length}
                </Text>
              )}
            </div>
          </DialogTitle>

          <DialogContent className="preview-content">
            {/* Preview Image with navigation arrows */}
            <div className="preview-image-container">
              {hasPrev && (
                <Button
                  className="preview-nav preview-nav-prev"
                  appearance="subtle"
                  icon={<ArrowLeftRegular />}
                  onClick={navigatePrev}
                  aria-label="Previous slide"
                />
              )}

              <div className="preview-image">
                {!imgError && asset.thumbnailPath ? (
                  <img
                    src={asset.thumbnailPath}
                    alt={`Slide ${asset.slideNumber}`}
                    onError={() => setImgError(true)}
                  />
                ) : (
                  <div className="preview-placeholder">
                    <ImageRegular style={{ fontSize: 48, color: "#A6B1BB" }} />
                    <Text size={300}>Preview not available</Text>
                  </div>
                )}
              </div>

              {hasNext && (
                <Button
                  className="preview-nav preview-nav-next"
                  appearance="subtle"
                  icon={<ArrowRightRegular />}
                  onClick={navigateNext}
                  aria-label="Next slide"
                />
              )}
            </div>

            {/* Slide Details */}
            <div className="preview-details">
              <div className="preview-detail-row">
                <Text size={200} weight="semibold">Category:</Text>
                <Badge appearance="filled" color="brand">
                  {asset.category}
                </Badge>
              </div>

              {asset.subcategory && (
                <div className="preview-detail-row">
                  <Text size={200} weight="semibold">Subcategory:</Text>
                  <Badge appearance="tint" color="informative">
                    {asset.subcategory}
                  </Badge>
                </div>
              )}

              <div className="preview-detail-row">
                <Text size={200} weight="semibold">Source:</Text>
                <Text size={200}>Slide #{asset.slideNumber}</Text>
              </div>

              {asset.tags.length > 0 && (
                <div className="preview-tags">
                  <Text size={200} weight="semibold">Tags:</Text>
                  <div className="preview-tag-list">
                    {asset.tags.map((tag, i) => (
                      <Badge
                        key={i}
                        appearance="outline"
                        size="small"
                        color="informative"
                      >
                        {tag}
                      </Badge>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </DialogContent>

          <DialogActions>
            <Button appearance="secondary" onClick={onClose}>
              Close
            </Button>
            <Button
              appearance="primary"
              icon={inserting ? <Spinner size="tiny" /> : <SlideAddRegular />}
              onClick={handleInsertSlide}
              disabled={inserting}
            >
              {inserting ? "Inserting…" : "Insert Slide"}
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
};

export default AssetPreview;
