import React, { useCallback, useState, useRef, useEffect, useMemo } from "react";
import {
  Card,
  Text,
  Button,
  Tooltip,
  Badge,
  Spinner,
  useToastController,
  Toast,
  ToastBody,
  ToastTitle,
} from "@fluentui/react-components";
import {
  SlideAddRegular,
  ArrowSortRegular,
} from "@fluentui/react-icons";
import { Asset } from "../../types";
import { insertSlide } from "../../services/officeApi";

interface AssetGridProps {
  assets: Asset[];
  onAssetClick: (asset: Asset) => void;
  toasterId: string;
}

type SortMode = "default" | "name-asc" | "name-desc" | "category";
const ITEMS_PER_PAGE = 30;
const SCROLL_THRESHOLD = 200; // px from bottom to trigger load more

const AssetGrid: React.FC<AssetGridProps> = ({
  assets,
  onAssetClick,
  toasterId,
}) => {
  const [displayCount, setDisplayCount] = useState(ITEMS_PER_PAGE);
  const [insertingId, setInsertingId] = useState<string | null>(null);
  const [sortMode, setSortMode] = useState<SortMode>("default");
  const { dispatchToast } = useToastController(toasterId);
  const scrollContainerRef = useRef<HTMLDivElement>(null);

  // Reset display count when assets change (e.g. new search/filter)
  useEffect(() => {
    setDisplayCount(ITEMS_PER_PAGE);
    // Scroll to top when results change
    if (scrollContainerRef.current) {
      scrollContainerRef.current.scrollTop = 0;
    }
  }, [assets]);

  // Sort assets
  const sortedAssets = useMemo(() => {
    const arr = [...assets];
    switch (sortMode) {
      case "name-asc":
        return arr.sort((a, b) => a.title.localeCompare(b.title));
      case "name-desc":
        return arr.sort((a, b) => b.title.localeCompare(a.title));
      case "category":
        return arr.sort((a, b) => a.category.localeCompare(b.category) || a.slideNumber - b.slideNumber);
      default:
        return arr; // default: slideNumber order (as indexed)
    }
  }, [assets, sortMode]);

  const visibleAssets = sortedAssets.slice(0, displayCount);
  const hasMore = displayCount < sortedAssets.length;

  // Infinite scroll: load more when user scrolls near bottom
  const handleScroll = useCallback(() => {
    const el = scrollContainerRef.current;
    if (!el || !hasMore) return;
    const { scrollTop, scrollHeight, clientHeight } = el;
    if (scrollHeight - scrollTop - clientHeight < SCROLL_THRESHOLD) {
      setDisplayCount((prev) => Math.min(prev + ITEMS_PER_PAGE, assets.length));
    }
  }, [hasMore, assets.length]);

  // Cycle sort mode
  const cycleSortMode = useCallback(() => {
    setSortMode((prev) => {
      const modes: SortMode[] = ["default", "name-asc", "name-desc", "category"];
      const idx = modes.indexOf(prev);
      return modes[(idx + 1) % modes.length];
    });
  }, []);

  const sortLabel = useMemo(() => {
    switch (sortMode) {
      case "name-asc": return "A → Z";
      case "name-desc": return "Z → A";
      case "category": return "Category";
      default: return "Default";
    }
  }, [sortMode]);

  const handleInsert = useCallback(
    async (e: React.MouseEvent, asset: Asset) => {
      e.stopPropagation(); // Don't open preview
      setInsertingId(asset.id);
      try {
        await insertSlide(asset);
        dispatchToast(
          <Toast>
            <ToastTitle>Inserted!</ToastTitle>
            <ToastBody>
              Slide #{asset.slideNumber} has been inserted as a native editable slide.
            </ToastBody>
          </Toast>,
          { intent: "success" }
        );
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
        setInsertingId(null);
      }
    },
    [dispatchToast]
  );

  // Thumbnail error handling with fallback
  const handleImageError = useCallback(
    (e: React.SyntheticEvent<HTMLImageElement>) => {
      const target = e.target as HTMLImageElement;
      target.style.display = "none";
      const parent = target.parentElement;
      if (parent) {
        parent.classList.add("thumbnail-placeholder");
      }
    },
    []
  );

  if (assets.length === 0) {
    return (
      <div className="asset-grid-empty">
        <div className="empty-icon">🔍</div>
        <Text size={300} weight="semibold">No assets found</Text>
        <Text size={200}>Try adjusting your search or category filter.</Text>
      </div>
    );
  }

  return (
    <div
      className="asset-grid-container"
      ref={scrollContainerRef}
      onScroll={handleScroll}
    >
      {/* Grid header with count and sort */}
      <div className="asset-grid-header">
        <Text size={200} weight="semibold">
          {assets.length} asset{assets.length !== 1 ? "s" : ""}
        </Text>
        <Tooltip content={`Sort: ${sortLabel}`} relationship="label">
          <Button
            appearance="transparent"
            size="small"
            icon={<ArrowSortRegular />}
            onClick={cycleSortMode}
            className="sort-button"
          >
            {sortLabel}
          </Button>
        </Tooltip>
      </div>

      {/* Thumbnail grid */}
      <div className="asset-grid">
        {visibleAssets.map((asset) => (
          <Card
            key={asset.id}
            className="asset-card"
            onClick={() => onAssetClick(asset)}
            appearance="subtle"
          >
            {/* Thumbnail */}
            <div className="asset-thumbnail">
              {asset.thumbnailPath ? (
                <img
                  src={asset.thumbnailPath}
                  alt={asset.title}
                  loading="lazy"
                  decoding="async"
                  onError={handleImageError}
                />
              ) : (
                <div className="thumbnail-placeholder">
                  <Text size={100}>{asset.category}</Text>
                </div>
              )}
            </div>

            {/* Info */}
            <div className="asset-info">
              <Tooltip content={asset.title} relationship="label">
                <Text size={100} className="asset-title" truncate wrap={false}>
                  {asset.title}
                </Text>
              </Tooltip>
              <div className="asset-meta">
                <Badge appearance="outline" size="small" color="informative">
                  {asset.category}
                </Badge>
                {asset.subcategory && (
                  <Badge appearance="outline" size="small" color="subtle">
                    {asset.subcategory}
                  </Badge>
                )}
              </div>
            </div>

            {/* Insert Button */}
            <div className="asset-actions">
              <Tooltip content="Insert as native slide" relationship="label">
                <Button
                  appearance="primary"
                  size="small"
                  icon={insertingId === asset.id ? <Spinner size="tiny" /> : <SlideAddRegular />}
                  onClick={(e) => handleInsert(e, asset)}
                  disabled={insertingId === asset.id}
                >
                  {insertingId === asset.id ? "Inserting…" : "Insert"}
                </Button>
              </Tooltip>
            </div>
          </Card>
        ))}
      </div>

      {/* Load More / Loading indicator */}
      {hasMore && (
        <div className="asset-grid-footer">
          <Text size={100} className="load-more-text">
            Showing {visibleAssets.length} of {assets.length} — scroll for more
          </Text>
        </div>
      )}
    </div>
  );
};

export default AssetGrid;
