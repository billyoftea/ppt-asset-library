import React, { useState, useEffect, useCallback, useMemo } from "react";
import {
  Spinner,
  Text,
  Toaster,
  useToastController,
  useId,
  Toast,
  ToastBody,
  ToastTitle,
} from "@fluentui/react-components";
import SearchBar from "./components/SearchBar";
import CategoryList from "./components/CategoryList";
import AssetGrid from "./components/AssetGrid";
import AssetPreview from "./components/AssetPreview";
import { AssetIndex, Asset, Category } from "../types";
import { loadAssetIndex } from "../services/assetManager";
import { preloadBundle } from "../services/officeApi";

const App: React.FC = () => {
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [assetIndex, setAssetIndex] = useState<AssetIndex | null>(null);
  const [searchQuery, setSearchQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState<string | null>(null);
  const [selectedAsset, setSelectedAsset] = useState<Asset | null>(null);
  const [filteredAssets, setFilteredAssets] = useState<Asset[]>([]);
  const [isOffline, setIsOffline] = useState(!navigator.onLine);

  const toasterId = useId("toaster");
  const { dispatchToast } = useToastController(toasterId);

  // 监听在线/离线状态
  useEffect(() => {
    const handleOnline = () => {
      setIsOffline(false);
      dispatchToast(
        <Toast>
          <ToastTitle>🌐 网络已恢复</ToastTitle>
          <ToastBody>资源将在后台自动更新缓存</ToastBody>
        </Toast>,
        { intent: "success", timeout: 3000 }
      );
    };
    const handleOffline = () => {
      setIsOffline(true);
      dispatchToast(
        <Toast>
          <ToastTitle>📴 已切换到离线模式</ToastTitle>
          <ToastBody>使用本地缓存，所有功能正常可用</ToastBody>
        </Toast>,
        { intent: "warning", timeout: 3000 }
      );
    };
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, [dispatchToast]);

  /**
   * 通知 Service Worker 预缓存所有缩略图
   */
  const triggerThumbnailCache = useCallback((index: AssetIndex) => {
    if ("serviceWorker" in navigator && navigator.serviceWorker.controller) {
      const thumbnailUrls = index.assets.map(
        (a) => `./assets/thumbnails/slide_${a.slideNumber}.png`
      );
      navigator.serviceWorker.controller.postMessage({
        type: "CACHE_THUMBNAILS",
        urls: thumbnailUrls,
      });
      console.log(`[App] Requested SW to cache ${thumbnailUrls.length} thumbnails`);
    }
  }, []);

  // Load asset index on mount
  useEffect(() => {
    const init = async () => {
      try {
        const index = await loadAssetIndex();
        setAssetIndex(index);
        setFilteredAssets(index.assets);
        // 预加载 bundle.pptx，加速首次插入
        preloadBundle().catch(() => {});
        // 通知 SW 缓存所有缩略图（后台静默执行）
        triggerThumbnailCache(index);
      } catch (err) {
        const msg = err instanceof Error ? err.message : "Unknown error";
        console.error("Failed to load asset index:", msg);
        setError(msg);
      } finally {
        setLoading(false);
      }
    };
    init();
  }, [triggerThumbnailCache]);

  // Search and filter logic
  useEffect(() => {
    if (!assetIndex) return;

    let results = assetIndex.assets;

    // Filter by category
    if (selectedCategory) {
      results = results.filter((a) => a.category === selectedCategory);
    }

    // Filter by search query (multi-word AND match)
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase().trim();
      const queryWords = query.split(/\s+/).filter(Boolean);

      results = results.filter((asset) => {
        const searchableText = [
          asset.title,
          asset.category,
          asset.subcategory,
          ...asset.tags,
          ...asset.keywords,
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();

        return queryWords.every((word) => searchableText.includes(word));
      });
    }

    setFilteredAssets(results);
  }, [searchQuery, selectedCategory, assetIndex]);

  const handleCategorySelect = useCallback((categoryName: string | null) => {
    setSelectedCategory(categoryName);
  }, []);

  const handleSearch = useCallback((query: string) => {
    setSearchQuery(query);
  }, []);

  const handleAssetClick = useCallback((asset: Asset) => {
    setSelectedAsset(asset);
  }, []);

  const handleClosePreview = useCallback(() => {
    setSelectedAsset(null);
  }, []);

  // Navigate within preview (prev/next)
  const handlePreviewNavigate = useCallback((asset: Asset) => {
    setSelectedAsset(asset);
  }, []);

  // Asset counts by category
  const assetCounts = useMemo(() => {
    if (!assetIndex) return {} as Record<string, number>;
    return assetIndex.categories.reduce(
      (acc, cat) => ({ ...acc, [cat.name]: cat.count }),
      {} as Record<string, number>
    );
  }, [assetIndex]);

  // Retry loading
  const handleRetry = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const index = await loadAssetIndex();
      setAssetIndex(index);
      setFilteredAssets(index.assets);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Unknown error");
    } finally {
      setLoading(false);
    }
  }, []);

  // --- Loading state ---
  if (loading) {
    return (
      <div className="app-loading">
        <Spinner size="large" label="Loading asset library..." />
      </div>
    );
  }

  // --- Error state ---
  if (error || !assetIndex) {
    return (
      <div className="app-error">
        <div className="error-icon">⚠️</div>
        <Text size={400} weight="semibold">Failed to load asset library</Text>
        <Text size={200}>{error || "Please check that assetIndex.json is available."}</Text>
        <button className="retry-button" onClick={handleRetry}>
          Try Again
        </button>
      </div>
    );
  }

  return (
    <div className="app-container">
      <Toaster toasterId={toasterId} position="bottom" />

      {/* Header */}
      <header className="app-header">
        <div className="app-header-row">
          <Text as="h1" size={500} weight="semibold">
            📦 Asset Library
          </Text>
        </div>
        <Text size={200} className="app-subtitle">
          {assetIndex.totalAssets} slides • {assetIndex.categories.length} categories
        </Text>
      </header>

      {/* Search Bar with result count */}
      <SearchBar
        onSearch={handleSearch}
        resultCount={filteredAssets.length}
        totalCount={assetIndex.totalAssets}
      />

      {/* Main Content */}
      <div className="app-content">
        {/* Category filter bar */}
        <CategoryList
          categories={assetIndex.categories}
          selectedCategory={selectedCategory}
          onSelectCategory={handleCategorySelect}
          assetCounts={assetCounts}
        />

        {/* Asset Grid with infinite scroll */}
        <AssetGrid
          assets={filteredAssets}
          onAssetClick={handleAssetClick}
          toasterId={toasterId}
        />
      </div>

      {/* Preview Modal with navigation */}
      {selectedAsset && (
        <AssetPreview
          asset={selectedAsset}
          onClose={handleClosePreview}
          toasterId={toasterId}
          allAssets={filteredAssets}
          onNavigate={handlePreviewNavigate}
        />
      )}
    </div>
  );
};

export default App;
