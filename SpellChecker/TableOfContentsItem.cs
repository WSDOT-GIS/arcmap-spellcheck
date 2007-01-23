using System;
using System.Collections;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Geometry;

namespace ArcMapSpellCheck {
	/// <summary>
	/// Represents an item that can be displayed in ArcMap's table of contents.
	/// </summary>
	/// <remarks>
	///		The purpose of this class is to have an object that can represent either a map or a layer,
	///		and get or set its Name property without the need to check what type of object it is and 
	///		then cast the object to that type before getting or setting the object's Name property.
	/// </remarks>
	internal class TableOfContentsItem {
		private object _tocObject;
		protected Descriptor _descriptor;

		#region Constructors
		/// <summary>
		/// Creates a new instance of <see cref="TableOfContentsItem"/>.
		/// </summary>
		public TableOfContentsItem(IMap map) {
			if (map == null) throw new ArgumentNullException("map");
			_tocObject = map;
			_descriptor = Descriptor.Map;
		}

		/// <summary>
		/// Creates a new instance of <see cref="TableOfContentsItem"/>.
		/// </summary>
		public TableOfContentsItem(ICompositeLayer layer) {
			if (layer == null) throw new ArgumentNullException("layer");
			_tocObject = layer;
			_descriptor = Descriptor.CompositeLayer;
		}

		/// <summary>
		/// Creates a new instance of <see cref="TableOfContentsItem"/>.
		/// </summary>
		public TableOfContentsItem(ILayer layer) {
			if (layer == null) throw new ArgumentNullException("layer");
			_tocObject = layer;
			_descriptor = Descriptor.Layer;
		}
		#endregion

		#region Properties
		/// <summary>
		/// The item in ArcMap's table of contents that this <see cref="TableOfContentsItem"/> object represents.
		/// </summary>
		public object Item {
			get { return _tocObject; }
		}

		/// <summary>
		/// Indicates the type of item.
		/// </summary>
		public Descriptor Descriptor {
			get { return _descriptor; }
		}

		/// <summary>
		/// Gets or sets the name of <see cref="TableOfContentsItem.Item"/>.
		/// </summary>
		public string Name {
			get {
				switch (_descriptor) {
					case Descriptor.Map:
						IMap map = _tocObject as IMap;
						return map.Name;
					case Descriptor.CompositeLayer:
					case Descriptor.Layer:
						ILayer layer = _tocObject as ILayer;
						return layer.Name;
					default:
						return null;
				}
			}
			set {
				switch (_descriptor) {
					case Descriptor.Map:
						IMap map = _tocObject as IMap;
						map.Name = value;
						break;
					case Descriptor.CompositeLayer:
					case Descriptor.Layer:
						ILayer layer = _tocObject as ILayer;
						layer.Name = value;
						break;
				}
			}
		}

		#endregion Properties

		/// <summary>
		/// Gets all maps and layers that are contained in a given <see cref="IMaps"/> object.
		/// </summary>
		/// <returns>
		///		An array of <see cref="TableOfContentsItem"/>s representing all of the maps and layers in <paramref name="maps"/>.
		///		If <paramref name="maps"/> is <see langword="null"/> or does not contain any maps or layers, an empty array will
		///		be returned.
		///	</returns>
		public static TableOfContentsItem[] GetAllTableOfContentsItems(IMaps maps) {
			ArrayList list = new ArrayList();
			if (maps != null) {

				// Loop through all maps.
				for (int mapIdx = 0; mapIdx < maps.Count; mapIdx++) {
					IMap currentMap = maps.get_Item(mapIdx);
					list.Add(new TableOfContentsItem(currentMap));
					
					// Create a list of all layers.
					ArrayList layersInMap = new ArrayList();
					for (int layerIdx = 0; layerIdx < currentMap.LayerCount; layerIdx++) {
						ILayer currentLayer = currentMap.get_Layer(layerIdx);
						layersInMap.AddRange(GetAllLayers(currentLayer));
					}
					
					foreach (ILayer currentLayer in layersInMap) {
						list.Add(new TableOfContentsItem(currentLayer));
					}
				}
			}
			
			return (TableOfContentsItem[])list.ToArray(typeof(TableOfContentsItem));
		}

		/// <summary>
		/// Gets all layers that are contained within a given <see cref="ILayer"/>.
		/// </summary>
		/// <param name="layer">An <see cref="ILayer"/>.</param>
		/// <returns>
		///		An <see cref="ArrayList"/> containing <paramref name="layer"/> and all of the layers that 
		///		<paramref name="layer"/> contains (and all of the layers that those child layers contain, etc.).
		///		If <paramref name="layer"/> is <see langword="null"/> an empty array will be returned.
		///	</returns>
		private static ArrayList GetAllLayers(ILayer layer) {
			ArrayList list = new ArrayList();
			if (layer != null) {
				list.Add(layer);
				ICompositeLayer compLayer = layer as ICompositeLayer;
				if (compLayer != null) {
					for (int layerIndex = 0; layerIndex < compLayer.Count; layerIndex++) {
						list.AddRange(GetAllLayers(compLayer.get_Layer(layerIndex)));
					}
				}
			}

			return list;
		}
	}
	

	/// <summary>
	/// Describes what kind of object is represented by a <see cref="TableOfContentsItem"/>.
	/// </summary>
	internal enum Descriptor {
		/// <summary>The object is not one of the valid types.</summary>
		Other = 0, 
		/// <summary>An <see cref="IMap"/> object.</summary>
		Map, 
		/// <summary>An <see cref="ICompositeLayer"/> object.</summary>
		CompositeLayer, 
		/// <summary>An <see cref="ILayer"/> object that is not an <see cref="ICompositeLayer"/></summary>
		Layer 
		//, Symbol
	}
}
