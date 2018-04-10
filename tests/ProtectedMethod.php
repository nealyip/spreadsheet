<?php
/**
 * Created by PhpStorm.
 * User: neal.yip
 * Date: 15/6/2017
 * Time: 16:42
 */

namespace Nealyip\Spreadsheet\Test;


trait ProtectedMethod
{

    /**
     * Test protected methods
     *
     * @param \stdClass|mixed $object
     * @param string          $method
     * @param array           $args
     *
     * @return mixed
     * @throws \ReflectionException
     */
    protected function _protectedMethod($object, $method, $args = [])
    {
        $reflection = new \ReflectionClass(get_class($object));
        $method     = $reflection->getMethod($method);
        $method->setAccessible(true);

        return $method->invokeArgs($object, $args);
    }

    /**
     * Get protected property
     *
     * @param \stdClass|mixed $object
     * @param                 $property
     *
     * @return mixed
     * @throws \ReflectionException
     */
    protected function _protectedProperty($object, $property)
    {
        $reflection = new \ReflectionClass(get_class($object));
        $rp         = $reflection->getProperty($property);
        $rp->setAccessible(true);
        return $rp->getValue($object);
    }

    /**
     * Set property
     *
     * @param \stdClass|mixed $object
     * @param string          $property
     * @param mixed           $value
     *
     * @throws \ReflectionException
     */
    protected function _setProperty($object, $property, $value)
    {

        $reflection = new \ReflectionClass(get_class($object));
        $rp         = $reflection->getProperty($property);
        $rp->setAccessible(true);
        $rp->setValue($object, $value);

    }
}