<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2011 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2011 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.6, 2011-02-27
 */


/**
 * PHPExcel_CachedObjectStorage_PHPTempBucket
 *
 * @category   PHPExcel
 * @package    PHPExcel_CachedObjectStorage
 * @copyright  Copyright (c) 2006 - 2011 PHPExcel (http://www.codeplex.com/PHPExcel)
 */
class PHPExcel_CachedObjectStorage_PHPTempBucket extends PHPExcel_CachedObjectStorage_CacheBase implements PHPExcel_CachedObjectStorage_ICache {

	private $_fileHandle = null;


	private $_memoryCacheSize = null;

	private $_bucketChanged = true;

	private $_bucket = array();
	private $_bucketMaxSize = null;

	public function isDataSet($pCoord)
	{
		if (isset($this->_bucket[$pCoord])) {
			return true;
		}

		return isset($this->_cellCache[$pCoord]);
	}	//	function isDataSet()


	private function _storeData() {
		if (!$this->_bucketChanged) {
			return;
		}

		foreach ($this->_bucket as $cell) {
			$cell->detach();
		}

		fseek($this->_fileHandle,0,SEEK_END);
		$offset = ftell($this->_fileHandle);
		fwrite($this->_fileHandle, igbinary_serialize($this->_bucket));

		foreach ($this->_bucket as $coord => $cell) {
			$this->_cellCache[$coord]	= array(
				'ptr' => $offset,
				'sz'  => ftell($this->_fileHandle) - $offset
				);
		}

		$this->_bucketChanged = false;
	}	//	function _storeData()


    /**
     *	Add or Update a cell in cache identified by coordinate address
     *
     *	@param	string			$pCoord		Coordinate address of the cell to update
     *	@param	PHPExcel_Cell	$cell		Cell to update
	 *	@return	void
     *	@throws	Exception
     */
	public function addCacheData($pCoord, PHPExcel_Cell $cell) {
		if (!isset($this->_bucket[$pCoord]) && count($this->_bucket) >= $this->_bucketMaxSize) {
			$this->_storeData();
			$this->_bucket = array();
		}

		$this->_bucket[$pCoord] = $cell;
		$this->_bucketChanged = true;
		if (!isset($this->_cellCache[$pCoord])) {
			$this->_cellCache[$pCoord] = null;
		}

		return $cell;
	}	//	function addCacheData()


    /**
     * Get cell at a specific coordinate
     *
     * @param 	string 			$pCoord		Coordinate of the cell
     * @throws 	Exception
     * @return 	PHPExcel_Cell 	Cell that was found, or null if not found
     */
	public function getCacheData($pCoord) {
		if (isset($this->_bucket[$pCoord])) {
			return $this->_bucket[$pCoord];
		}
		$this->_storeData();

		//	Check if the entry that has been requested actually exists
		if (!isset($this->_cellCache[$pCoord])) {
			//	Return null if requested entry doesn't exist in cache
			return null;
		}

		//	Set current entry to the requested entry
		fseek($this->_fileHandle,$this->_cellCache[$pCoord]['ptr']);
		$this->_bucket = igbinary_unserialize(fread($this->_fileHandle,$this->_cellCache[$pCoord]['sz']));
		//	Re-attach the parent worksheet
		foreach ($this->_bucket as $cell) {
			$cell->attach($this->_parent);
		}

		//	Return requested entry
		return $this->_bucket[$pCoord];
	}	//	function getCacheData()


	/**
	 *	Clone the cell collection
	 *
	 *	@return	void
	 */
	public function copyCellCollection(PHPExcel_Worksheet $parent) {
		parent::copyCellCollection($parent);
		//	Open a new stream for the cell cache data
		$newFileHandle = fopen('php://temp/maxmemory:'.$this->_memoryCacheSize,'a+');
		//	Copy the existing cell cache data to the new stream
		fseek($this->_fileHandle,0);
		while (!feof($this->_fileHandle)) {
			fwrite($newFileHandle,fread($this->_fileHandle, 1024));
		}
		$this->_fileHandle = $newFileHandle;
	}	//	function copyCellCollection()


	public function deleteCacheData($pCoord) {
		if (isset($this->_bucket[$pCoord])) {
			$this->_bucket[$pCoord]->detach();
			unset($this->_bucket[$pCoord]);
		}

		if (isset($this->_cellCache[$pCoord])) {
			unset($this->_cellCache[$pCoord]);
		}
	}	//	function deleteCacheData()


	public function unsetWorksheetCells() {
		foreach ($this->_bucket as $cell) {
			$cell->detach();
		}
		$this->_bucket = array();

		$this->_cellCache = array();

		//	detach ourself from the worksheet, so that it can then delete this object successfully
		$this->_parent = null;

		//	Close down the php://temp file
		$this->__destruct();
	}	//	function unsetWorksheetCells()


	public function __construct(PHPExcel_Worksheet $parent, $arguments) {
		$this->_memoryCacheSize	= (isset($arguments['memoryCacheSize'])) ? $arguments['memoryCacheSize'] : '1MB';
		$this->_bucketMaxSize = (isset($arguments['bucketMaxSize'])) ? $arguments['bucketMaxSize'] : 20;

		parent::__construct($parent);
		if (is_null($this->_fileHandle)) {
			$this->_fileHandle = fopen('php://temp/maxmemory:'.$this->_memoryCacheSize,'a+');
		}
	}	//	function __construct()


	public function __destruct() {
		if (!is_null($this->_fileHandle)) {
			fclose($this->_fileHandle);
		}
		$this->_fileHandle = null;
	}	//	function __destruct()

}
